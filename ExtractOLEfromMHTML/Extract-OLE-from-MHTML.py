import zlib
import base64
import re
import sys

def get_zlib_data(data):
    if not re.match("ActiveMime",data):
        return ""
    found = re.search('\x78\x9c',data,re.MULTILINE)
    if found:
        return data[found.start():len(data)]

def get_ole_from_mhtml(filepath):
    fh = file(filepath,"rb")
    fdata = fh.read()
    fh.close()
    
    found = re.search("^Content-Location:\x20file:///[^\n]{0,999}?editdata\.mso.*?\r\n\r\n^((?:[A-Za-z0-9+/\r\n]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?)\r\n",fdata,re.DOTALL|re.MULTILINE|re.IGNORECASE)
    if found:
        ole = found.group(1)
    return ole
    

if __name__ == '__main__':
    if len(sys.argv) > 1:
        print sys.argv[1]
        g = get_ole_from_mhtml(sys.argv[1])
        activemimestream = base64.b64decode(g)
        zlibdata = get_zlib_data(activemimestream)
        olestream = zlib.decompress(zlibdata)
        fh = file(sys.argv[1]+".ole","wb")
        fh.write(olestream)
        fh.close()