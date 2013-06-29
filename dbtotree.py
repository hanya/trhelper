
import os
import os.path
import argparse
from xml.sax.saxutils import escape
try:
    from cStringIO import StringIO
except:
    from io import StringIO


def read_db(path):
    """ Reads key, value pairs from ht, db or key file. """
    if not os.path.exists(path): raise Exception("File not found: " + path)
    pairs = []
    apairs = pairs.append
    
    def _read_length(f):
        v = []
        while True:
            c = f.read(1)
            _c = ord(c)
            if 0x30 <= _c <= 0x39 \
                or 0x41 <= _c <= 0x46 \
                or 0x61 <= _c <= 0x66:
                v.append(c)
            else:
                return int("".join(v), 16)
    try:
        with open(path) as f:
            while True:
                n = _read_length(f) # read key length
                key = f.read(n)
                f.seek(1, 1)        # move to next length for value
                n = _read_length(f) # read value length
                value = f.read(n)
                apairs((key, value))
                f.seek(1, 1)        # skip next \n
    except:
        pass
    return dict(pairs)


def split_db_values(d):
    """ Value of db file contains combined three values. 
        Splits all values. """
    entries = []
    aentries = entries.append
    for key, value in d.items():
        # split into three values
        n = ord(value[0])
        m = ord(value[n+1])
        #o = ord(value[n+2+m])
        aentries((key, (value[1:n+1], value[n+2:n+1+m], value[n+2+m+1:])))
    return dict(entries)


tree_header = """<?xml version="1.0" encoding="UTF-8"?>
<tree_view version="00-7-2013"><help_section application="all" id="123223" title="All">\n"""
tree_footer = """</help_section></tree_view>\n"""



class MergedTree(object):
    """ Generates tree that contains list of pages. """
    def __init__(self, base_path, lang, modules, module_names):
        self.base_path = base_path
        self.lang = lang
        self.modules = modules
        self.module_names = module_names
        if len(modules) != len(module_names):
            raise Exception("Number of modules do not match with their names")
    
    def store(self, out_base_path):
        """ Read db file and generates tree and cfg files for specified language. """
        _modules = {}
        for module in self.modules:
            _modules[module] = self._read_db(module, "db")
        
        # extract text/shared/ pages from each modules
        shared = {}
        for module, pages in _modules.items():
            candidate = []
            for key, value in pages.items():
                if key.startswith("text/shared/"):
                    #del pages[key]
                    candidate.append(key)
                    # merge value
                    try:
                        _value = shared[key]
                        value[1].update(_value[1])
                    except:
                        shared[key] = (value[0], set(value[1])) # title, value
            for name in candidate:
                del pages[name]
        
        # add shared to list of modules
        _modules["shared"] = shared # replace existing shared
        
        io = StringIO()
        w = io.write
        w(tree_header)
        
        _escape = escape
        for module, module_name in zip(self.modules, self.module_names):
            # sort in page path
            w("<node id=\"{ID}\" title=\"{MODULE_NAME}\">\n".format(
                    ID=hash(module_name), MODULE_NAME=module_name + " - Sorted in page path"))
            pages = _modules[module]
            for page, value in sorted(pages.items()):
                w("<topic id=\"{MODULE}/{PAGE}\">{TITLE}</topic>\n".format(
                    MODULE=module, PAGE=page, TITLE=_escape(value[0])))
            w("</node>\n")
            
            # sort in page title
            w("<node id=\"{ID}\" title=\"{MODULE_NAME}\">\n".format(
                    ID=hash(module_name)+1, MODULE_NAME=module_name + " - Sorted in page title"))
            entries = [(value[0], page) for page, value in pages.items()]
            for value, page in sorted(entries):
                w("<topic id=\"{MODULE}/{PAGE}\">{TITLE}</topic>\n".format(
                    MODULE=module, PAGE=page, TITLE=_escape(value)))
            w("</node>\n")
        
        w(tree_footer)
        
        _base_dir = os.path.join(out_base_path, self.lang)
        try:
            os.makedirs(_base_dir)
        except:
            pass
        
        with open(os.path.join(_base_dir, "all.tree"), "w") as f:
            f.write(io.getvalue())
            f.flush()
        io.close()
        # write cfg
        with open(os.path.join(_base_dir, "all.cfg"), "w") as f:
            f.write("""
Title=%PRODUCTNAME All
Language={LANG}
Order=8
Start=text%2Fshared%2Fmain0000.xhp
Heading=headingheading
Program=ALL 
07.07.04 00:00:00
""".format(LANG=self.lang))
    
    def _read_db(self, module, _type):
        entries = read_db(os.path.join(self.base_path, self.lang, module) + "." + _type)
        data = split_db_values(entries)
        del entries
        pages = {}
        for k, (path, _, title) in data.items():
            parts = path.split("#", 1)
            page_path = parts[0]
            anchor = parts[1] if len(parts) >= 2 else ""
            try:
                pages[page_path][1].append(anchor)
            except:
                pages[page_path] = (title, [anchor])
        return pages


def parse_arguments():
    desc = """Reads db files and generates complete tree."""
    #  "python dbtotree.py -i path_to_help_dir -o output_dir"
    parser = argparse.ArgumentParser(description=desc)
    parser.add_argument("-i", help="input base path that contains help files", 
                            dest="help_path", metavar="INPUT_DIR", required=True)
    parser.add_argument("-o", help="output base path", 
                            dest="out_path", metavar="OUTPUT_DIR", required=True)
    
    return parser.parse_args()


def main():
    args = parse_arguments()
    help_path = args.help_path
    out_path = args.out_path
    
    modules = ["shared", "swriter", "scalc", "simpress", "sdraw", "sdatabase", "smath", "schart", "sbasic"]
    module_names = ["Shared", "Writer", "Calc", "Impress", "Draw", "Base", "Math", "Chart", "Basic"]
    
    for lang in os.listdir(help_path):
        if os.path.isdir(os.path.join(help_path, lang)):
            MergedTree(help_path, lang, modules, module_names).store(out_path)


if __name__ == "__main__":
    main()
