#!/opt/openoffice.org3/program/python
import os
import sys
import time
import socket
import shutil
import tempfile
import subprocess
from optparse import OptionParser

import uno
import unohelper
from com.sun.star.connection import NoConnectException
from com.sun.star.beans import PropertyValue
from com.sun.star.io import XOutputStream

OPENOFFICE_BIN = '/opt/openoffice.org3/program/swriter'

FILTER_MAP = {
    'doc': 'MS Word 97',
    'docx': 'MS Word 2007 XML',
    'odt': 'writer8',
    'pdf': 'writer_pdf_Export',
    'rtf': 'Rich Text Format',
    'txt': 'Text (encoded)',
    'html': 'HTML (StarWriter)',
}


class OutputStream(unohelper.Base, XOutputStream):
    def __init__(self, descriptor=None):
        self.descriptor = descriptor
        self.closed = 0

    def closeOutput(self):
        self.closed = 1
        if not self.descriptor.isatty:
            self.descriptor.close()

    def writeBytes(self, seq):
        self.descriptor.write(seq.value)

    def flush(self):
        pass


def to_properties(d):
    return tuple(PropertyValue(key, 0, value, 0) for key, value in d.iteritems())


def start_openoffice(home_dir, port):
    args = [OPENOFFICE_BIN,
            '-accept=socket,host=localhost,port=%d;urp;StarOffice.ServiceManager' % port,
            '-userid="%s"' % home_dir, '-norestore', '-nofirststartwizard', '-nologo',
            '-nocrashreport', '-nodefault', '-quickstart', '-norestart', '-nolockcheck',
            '-headless', '-invisible']

    custom_env = os.environ.copy()
    custom_env['HOME'] = home_dir
    try:
        popen = subprocess.Popen(args, env=custom_env)
        pid = popen.pid
    except Exception, e:
        print >> sys.stderr, 'Failed to start OpenOffice on port %d: %s' % (port, e.message)
        exit(1)

    if pid <= 0:
        print >> sys.stderr, 'Failed to start OpenOffice on port %d' % port
        exit(1)

    connection_params = 'uno:socket,host=%s,port=%s;urp;' \
                        'StarOffice.ComponentContext' % ('localhost', port)

    uno_context = None
    for times in range(20):
        context = uno.getComponentContext()
        resolver = context.ServiceManager.createInstanceWithContext(
                "com.sun.star.bridge.UnoUrlResolver", context)
        try:
            uno_context = resolver.resolve(connection_params)
            break
        except NoConnectException:
            time.sleep(1)
    if not uno_context:
        exit(1)

    desktop = uno_context.ServiceManager.createInstanceWithContext(
            'com.sun.star.frame.Desktop', context)
    return popen, context, desktop


def get_free_port():
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.bind(('0.0.0.0', 0))
    _, port = s.getsockname()
    s.close()
    return port


def convert(source, target, target_format):
    home_dir = tempfile.mkdtemp()
    errcode = None
    popen = None
    try:
        port = get_free_port()
        popen, context, desktop = start_openoffice(home_dir, port)
        
        input_stream = context.ServiceManager.createInstanceWithContext(
                'com.sun.star.io.SequenceInputStream', context)
        input_stream.initialize((uno.ByteSequence(source.read()),))

        doc = desktop.loadComponentFromURL('private:stream', '_blank', 0, to_properties({
           'InputStream': input_stream,
        }))

        try:
            doc.refresh()
        except AttributeError:
            pass
        try:
            doc.storeToURL('private:stream', to_properties({
                'FilterName': FILTER_MAP[target_format],
                'OutputStream': OutputStream(target)
            }))
        finally:
            doc.close(True)
    
        try:
            desktop.terminate()
        except:
            pass

        if popen.returncode is None:
            errcode = popen.wait()
    except Exception as e:
        print >> sys.stderr, e
        errcode = 1
    finally:
        if errcode is None and popen and popen.returncode is None:
            try:
                errcode = popen.kill()
            except:
                pass
        time.sleep(0.5)
        shutil.rmtree(home_dir)
        exit(errcode)


def main():
    usage = '''usage: %prog [options] source target
    `source` is the source file path or '-' for stdin;
    `target` is the target file path or '-' for stdout.
    If `target` is '-' or has no extension, --format option must be specified.
    '''
    parser = OptionParser(usage=usage)
    parser.add_option('-f', '--format', dest='target_format',
            help='Target format of the file', metavar='FORMAT')
    
    try:
        (options, (source_file, target_file,)) = parser.parse_args()
    except ValueError:
        parser.error('Two positional arguments are required.')

    target_format = options.target_format

    _, target_file_ext = os.path.splitext(target_file)
    target_file_ext = target_file_ext.replace('.', '')

    if not target_format and target_file_ext in FILTER_MAP.keys():
        target_format = target_file_ext

    if target_format is None:
        parser.error('You must specify target format.')

    source = sys.stdin if source_file == '-' else open(source_file, 'r')
    target = sys.stdout if target_file == '-' else open(target_file, 'w')
    convert(source, target, target_format)


if __name__ == '__main__':
    main()
