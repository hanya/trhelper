
import re
import time
import threading

import unohelper
from com.sun.star.awt import XActionListener
from com.sun.star.lang import XServiceInfo
from com.sun.star.task import XJobExecutor

from com.sun.star.awt import WindowDescriptor, Rectangle
from com.sun.star.awt.WindowClass import SIMPLE
from com.sun.star.awt.WindowAttribute import SHOW
from com.sun.star.awt.PosSize import WIDTH as PS_WIDTH
from com.sun.star.beans import PropertyValue
from com.sun.star.beans.PropertyState import DEFAULT_VALUE


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


class HelpHacker(unohelper.Base, ServiceInfo, XJobExecutor):
    """ Adds some function to help viewer. """
    
    IMPLE_NAME = "foo.bar.hoge.help.HelpHacker"
    SERVICE_NAMES = IMPLE_NAME,
    
    def __init__(self, ctx, *args):
        self.ctx = ctx
    
    # XJobExecutor
    def trigger(self, arg):
        try:
            if arg == "mode=pootle":
                mode = "pootle"
            elif arg == "mode=omegat":
                mode = "omegat"
            if mode:
                self.hack_help_viewer(mode)
        except Exception as e:
            message(self.ctx, self.get_current_doc(), str(e), "Error")
            import traceback
            traceback.print_exc()
    
    def hack_help_viewer(self, mode):
        # find frame for help viewer, inner frame.
        frame, state = self.find_help_view()
        args = (PropertyValue(),)
        args[0].Name = "Mode"
        args[0].Value = mode
        if state:
            self.add_en_button(frame)
            self.add_menu_button(frame)
            self.register_search_interception(args)
        else:
            # wait for the cunstruction delay
            threading.Timer(4, self.add_en_button, (frame,)).start()
            threading.Timer(5, self.add_menu_button, (frame,)).start()
            threading.Timer(4, self.register_search_interception, (args,)).start()
    
    def register_search_interception(self, mode_args):
        interception = self.create_service(
            "foo.bar.hoge.help.FooBarSearchDispatchInterceptorForHelpViewer", mode_args)
        interception.register()
    
    def find_help_view(self):
        def _find():
            return self.find_frame(self.get_desktop(), "OFFICE_HELP_TASK")
        frame = _find()
        if not frame:
            # call .uno:Help to open help viewer
            thread = threading.Thread(target=self.open_help_viewer)
            thread.start()
            thread.join(2)
            time.sleep(2)
            frame = _find()
            if not frame:
                raise Exception("Please open help viewer.") # open my self?
            return frame, False # not yet created
        return frame, True
    
    def open_help_viewer(self):
        self.execute_command(self.get_desktop(), ".uno:HelpIndex")
    
    def create_service(self, name, args=None):
        if args:
            return self.ctx.getServiceManager().\
                createInstanceWithArgumentsAndContext(name, args, self.ctx)
        return self.ctx.getServiceManager().createInstanceWithContext(name, self.ctx)
    
    def get_current_doc(self):
        return self.get_desktop().getCurrentComponent()
    
    def get_desktop(self):
        return self.create_service("com.sun.star.frame.Desktop")
    
    def open_url(self, url):
        arg = PropertyValue("FilterName", 0, "writer_web_HTML_help", DEFAULT_VALUE)
        self.get_desktop().loadComponentFromURL(url, "_blank", 0, (arg,))
    
    def open_help_url(self, url):
        self.open_url(url)
    
    def open_en_help(self, url):
        _url = re.sub("Language=([a-z]+-[A-Z]+|[a-z]+)", "Language=en", url)
        self.open_help_url(_url)
    
    def find_frame(self, container, name):
        """ Find the named frame inside the container. """
        return container.findFrame(name, 4)
    
    
    def add_en_button(self, frame):
        # add en button to open en-US help files
        inner_frame = self.find_frame(frame, "OFFICE_HELP")
        window = self.get_toolbar_window(frame)
        
        self.add_toolbar_button(window, "en", 
                self.ActionListener(self.Foo(self, inner_frame)))
    
    def add_menu_button(self, frame):
        inner_frame = self.find_frame(frame, "OFFICE_HELP")
        window = self.get_toolbar_window(frame)
        
        self.add_toolbar_button(window, "v", 
                self.ActionListener(self.Menu(self, inner_frame)))
    
    def add_toolbar_button(self, window, label, listener):
        ps = window.getPosSize()
        window.setPosSize(0, 0, ps.Width + 33, 0, PS_WIDTH)
        
        desc = WindowDescriptor(SIMPLE, "pushbutton", window, 0, 
                Rectangle(ps.Width + 3, 0, 30, 28), SHOW)
        btn = window.getToolkit().createWindow(desc)
        btn.Label = label
        btn.addActionListener(listener)
        return btn
    
    def get_toolbar_window(self, frame):
        # check the side window is opened
        comp_win = frame.getComponentWindow()
        acc_ctx = comp_win.getAccessibleContext()
        count = acc_ctx.getAccessibleChildCount()
        if count == 1:
            acc_child = acc_ctx.getAccessibleChild(0)
        elif count == 2:
            # side window is opened
            acc_child = acc_ctx.getAccessibleChild(1)
        # error?
        
        windows = acc_child.getWindows()
        return windows[0]
    
    def do_something(self, cmd):
        try:
            getattr(self, "do_" + cmd)()
        except:
            pass
    
    #def do_saveas(self):
    
    def execute_command(self, frame, cmd):
        dh = self.create_service("com.sun.star.frame.DispatchHelper")
        dh.executeDispatch(frame, cmd, "_self", 0, ())
    
    class Base(object):
        def __init__(self, act, frame):
            self.act = act
            self.frame = frame
    
    class Foo(Base):
        def action(self, source, cmd):
            if self.frame and self.act:
                url = self.frame.getController().getModel().getURL()
                self.act.open_en_help(url)
    
    class Menu(object):
        def __init__(self, act, frame):
            self.act = act
            self.frame = frame
            self.use_point = check_method_parameter(self.act.ctx, 
                "com.sun.star.awt.XPopupMenu", "execute", 1, "com.sun.star.awt.Point")
        
        def action(self, source, cmd):
            menu = self.act.create_service("com.sun.star.awt.PopupMenu")
            menu.insertItem(101, "Save As", 0, 0)
            menu.setCommand(101, "saveas")
            y = source.getPosSize().Height
            pos = Point(0, y) if self.use_point else Rectangle(0, y, 0, 0)
            n = menu.execute(source, pos, 0)
            if 0 < n < 100:
                self.act.do_something(menu.getCommand(n))
            elif n == 101:
                self.act.execute_command(self.frame, ".uno:SaveAs")
    
    class ActionListener(unohelper.Base, XActionListener):
        
        def __init__(self, act):
            self.act = act
        
        def disposing(self, ev):
            self.act = None
        
        def actionPerformed(self, ev):
            try:
                if self.act: self.act.action(ev.Source, ev.ActionCommand)
            except Exception as e:
                print(e)


def check_method_parameter(ctx, interface_name, method_name, param_index, param_type):
    """ Check the method has specific type parameter at the specific position. """
    cr = ctx.getServiceManager().createInstanceWithContext("com.sun.star.reflection.CoreReflection", ctx)
    try:
        idl = cr.forName(interface_name)
        m = idl.getMethod(method_name)
        if m:
            info = m.getParameterInfos()[param_index]
            return info.aType.getName() == param_type
    except:
        pass
    return False


from com.sun.star.awt import Rectangle

def message(ctx, doc, message, title):
    """ Shows message. """
    older = check_method_parameter(ctx, "com.sun.star.awt.XMessageBoxFactory", 
        "createMessageBox", 1, "com.sun.star.awt.Rectangle")
    
    window = doc.getCurrentController().getFrame().getContainerWindow()
    toolkit = window.getToolkit()
    if older:
        msgbox = toolkit.createMessageBox(window, Rectangle(), "messbox", 1, title, message)
    else:
        from com.sun.star.awt.MessageBoxType import MESSAGEBOX
        msgbox = toolkit.createMessageBox(window, MESSAGEBOX, 1, title, message)
    n = msgbox.execute()
    msgbox.dispose()
    return n


#if __name__ == "ooo_script_framework":
#def test(*args):
#    HelpHacker(XSCRIPTCONTEXT.getComponentContext()).trigger("")

g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(*HelpHacker.get_info())
