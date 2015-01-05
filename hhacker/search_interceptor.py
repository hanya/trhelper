
import webbrowser
try:
    from urllib2 import urlopen
    from urlparse import urlparse, parse_qs
except:
    from urllib.request import urlopen
    from urllib.parse import urlparse, parse_qs
import unohelper

from com.sun.star.frame import XDispatchProviderInterceptor, XDispatch, \
    XControlNotificationListener
from com.sun.star.lang import XInitialization, XServiceInfo
from com.sun.star.awt import XContainerWindowEventHandler
from com.sun.star.beans import PropertyValue


class ServiceInfo(XServiceInfo):
    
    # XServiceInfo
    def getImplementationName(self):
        return self.IMPLE_NAME
    
    def getSupportedServiceNames(self):
        return self.SERVICE_NAMES
    
    def supportsService(self, name):
        return name in self.SERVICE_NAMES
    
    @classmethod
    def get_info(klass):
        return klass, klass.IMPLE_NAME, klass.SERVICE_NAMES


class OptionsPageHandler(unohelper.Base, 
        XContainerWindowEventHandler, ServiceInfo):
    
    IMPLE_NAME = "foo.bar.hoge.help.FooBarSearchOptionsPageHandler"
    SERVICE_NAMES = IMPLE_NAME,
    
    def __init__(self, ctx, *args):
        self.ctx = ctx
        self.dialog = None
    
    # XContainerWindowEventHandler
    def getSupportedMethodNames(self):
        return ("external_event", )
    
    def callHandlerMethod(self, window, ev, name):
        if name == "external_event":
            self.handle(window, ev)
    
    def handle(self, dialog, ev):
        self.dialog = dialog
        name = self.dialog.getModel().Name
        if ev == "ok":
            getattr(self, "confirm_" + name)()
        elif ev == "back":
            getattr(self, "init_" + name)()
        elif ev == "initialize":
            getattr(self, "init_" + name)(first_time=True)
        return True
    
    def get(self, name):
        return self.dialog.getControl(name)
    
    def get_text(self, name):
        return self.get(name).getModel().Text
    
    def set_text(self, name, text):
        if text is None:
            text = ""
        self.get(name).getModel().Text = text
    
    def init_Options(self, first_time=False):
        try:
            d = Config(self.ctx).get_config_map()
            for key, value in d.items():
                self.set_text("edit{}".format(key), value)
        except Exception as e:
            print(e)
    
    def confirm_Options(self):
        names = self.dialog.getModel().getElementNames()
        d = {name[4:]: self.get_text(name) for name in names if name.startswith("edit")}
        Config(self.ctx).set_from_map(d)


class Config:
    """ Configuration reader-writer. """
    NODE_NAME = "/foo.bar.hoge.help.FooBarSearch/Settings"
    
    def __init__(self, ctx, modifiable=False):
        self.ctx = ctx
        self.config = self.get_config(self.NODE_NAME, modifiable)
    
    def create_service(self, name):
        return self.ctx.getServiceManager().createInstanceWithContext(name, self.ctx)
    
    def get_config(self, nodepath, modifiable=False):
        """ Get configuration node. """
        cp = self.create_service("com.sun.star.configuration.ConfigurationProvider")
        node = PropertyValue("nodepath", -1, nodepath, 0)
        if modifiable:
            name = "com.sun.star.configuration.ConfigurationUpdateAccess"
        else:
            name = "com.sun.star.configuration.ConfigurationAccess"
        return cp.createInstanceWithArguments(name, (node,))
    
    def get(self, name):
        return self.config.getPropertyValue(name)
    
    def get_config_map(self):
        config = self.config
        return {name: config.getPropertyValue(name) for name in config.getElementNames()}
    
    def set_from_map(self, d):
        config = self.config
        for key, value in d.items():
            config.setPropertyValue(key, value)
        config.commitChanges()


class Dispatcher(unohelper.Base, XDispatch, XControlNotificationListener):
    """ Dispatcher for .uno:FooBarSearch. 
    
        This class constructs search URL and access to it.
        # text/shared/01/online_update.xhp%23hd_id315256.help.text
    """
    
    def __init__(self, config, mode="pootle"):
        self.mode = mode
        self.pootle_url = config["PootleURL"]
        self.omegat_url = config["OmegaTURL"]
        self.help_port = config["HelpPort"]
        self.ui_port = config["UIPort"]
        self.project = config["PootleProjectName"]
    
    # XDispatch
    def dispatch(self, url, args):
        try:
            r = urlparse(url.Complete)
            if r.scheme == ".uno" and r.path == "FooBarSearch":
                qs = parse_qs(r.query)
                if "keyword" in qs:
                    mode = self.mode
                    if mode == "pootle":
                        self.search_in_pootle(r, qs)
                    elif mode == "omegat":
                        self.search_in_omegat(r, qs)
        except Exception as e:
            print(e)
    
    def search_in_pootle(self, r, qs):
        keyword = qs["keyword"][0] + "%23" + r.fragment
        category = qs["category"][0] if "category" in qs else "help"
        if category == "ui": category = ""
        if "language" in qs:
            # last part only
            _keyword = keyword.split("/")[-1]
            
            url = self.pootle_url.format(
                    LANG=qs["language"][0], PROJECT=self.project, CATEGORY=category) + _keyword
            webbrowser.open(url)
    
    def search_in_omegat(self, r, qs):
        keyword = qs["keyword"][0] + "%23" + r.fragment
        category = qs["category"][0] if "category" in qs else "help"
        port = self.help_port if category == "help" else self.ui_port
        
        _keyword = keyword.lstrip("/")
        url = self.omegat_url.format(PORT=port) + _keyword
        try:
            f = urlopen(url)
            f.close()
        except:
            pass
    
    def addStatusListener(self, xControl, url): pass
    
    def removeStatusListener(self, xControl, url): pass
    
    # XControlNotificationListener
    def controlEvent(self, ev): pass


class FooBarSearchDispatchInterceptor(unohelper.Base, ServiceInfo, XDispatchProviderInterceptor, XInitialization):
    """ This dispatch interceptor allows to execute custom .uno: command. 
        
        Through hyperlinks on documents, only uno commands are executed correctly. 
        Custom protocol did not work. 
        Commands should be: .uno:FooBarSearch?language=TARGET&category=(ui|help)&keyword=KEYWORD
        @param language target locale to search for in ISO code.
        @param category ui or help to search for
        @param keyword for search will be processed by the dispatcher
    """
    
    IMPLE_NAME = "foo.bar.hoge.help.FooBarSearchDispatchInterceptor"
    SERVICE_NAMES = IMPLE_NAME,
    
    def __init__(self, ctx, *args):
        self.ctx = ctx
        self.mode = "pootle"
        self.slave = None
        self.master = None
        self.dispatcher = None
        self.initialize(args)
    
    def create_service(self, name):
        return self.ctx.getServiceManager().createInstanceWithContext(name, self.ctx)
    
    def get_current_doc(self):
        return self.get_desktop().getCurrentComponent()
    
    def get_desktop(self):
        return self.create_service("com.sun.star.frame.Desktop")
    
    def find_frame(self, container, name):
        """ Find the named frame inside the container. """
        return container.findFrame(name, 4)
    
    def get_target_frame(self):
        # current component
        return self.get_deskop().getActiveFrame()
    
    def register(self):
        frame = self.get_target_frame()
        frame.registerDispatchProviderInterceptor(self)
    
    def release(self):
        frame = self.get_target_frame()
        frame.releaseDispatchProviderInterceptor(self)
    
    # XInitialization
    def initialize(self, args):
        for arg in args:
            if arg.Name == "Mode":
                self.mode = arg.Value
        if not self.dispatcher:
            self.dispatcher = Dispatcher(Config(self.ctx).get_config_map(), self.mode)
    
    # XDispatchProvider
    def queryDispatch(self, url, name, flag):
        if url.Complete.startswith(".uno:FooBarSearch?"):
            try:
                return self.dispatcher
            except Exception as e:
                print(e)
                return None
        try:
            if self.slave:
                return self.slave.queryDispatch(url, name, flag)
        except Exception as e:
            print(e)
        return None
    
    def queryDispatches(self, descs):
        return tuple(
            [self.queryDispatch(desc.FeatureURL, desc.FrameName, desc.SearchFlags) 
                for desc in descs])
    
    # XDispatchProviderInterceptor
    def getSlaveDispatchProvider(self):
        return self.slave
    
    def setSlaveDispatchProvider(self, provider):
        self.slave = provider
    
    def getMasterDispatchProvider(self):
        return self.master
    
    def setMasterDispatchProvider(self, supplier):
        self.master = supplier


class ForHelpViewer(FooBarSearchDispatchInterceptor):
    
    IMPLE_NAME = "foo.bar.hoge.help.FooBarSearchDispatchInterceptorForHelpViewer"
    SERVICE_NAMES = IMPLE_NAME,
    
    def find_help_view(self):
        def _find():
            return self.find_frame(self.get_desktop(), "OFFICE_HELP_TASK")
        return _find()

    def get_inner_frame(self):
        frame = self.find_help_view()
        if not frame:
            raise Exception("Open help viewer.")
        return self.find_frame(frame, "OFFICE_HELP")

    def get_target_frame(self):
        # inner frame of the help viewer
        return self.get_inner_frame()


interceptor = None

def get_interceptor(ctx):
    global interceptor
    if not interceptor:
        interceptor = FooBarSearchDispatchInterceptor(ctx)
    return interceptor


def register_help(*args):
    ForHelpViewer(XSCRIPTCONTEXT.getComponentContext()).register()


def register_current(*args):
    interceptor = get_interceptor(XSCRIPTCONTEXT.getComponentContext())
    interceptor.register()

def release(*args):
    interceptor = get_interceptor(XSCRIPTCONTEXT.getComponentContext())
    interceptor.release()

g_exportedScripts = register_current, register_help


g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    *FooBarSearchDispatchInterceptor.get_info())
g_ImplementationHelper.addImplementation(
    *ForHelpViewer.get_info())
g_ImplementationHelper.addImplementation(
    *OptionsPageHandler.get_info())
