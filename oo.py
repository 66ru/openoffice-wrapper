#! /usr/bin/env python
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


OPENOFFICE_BIN = '/usr/bin/libreoffice'
FILTER_MAP = {
    'doc': 'MS Word 97',
    'docx': 'MS Word 2007 XML',
    'odt': 'writer8',
    'pdf': 'writer_pdf_Export',
    'rtf': 'Rich Text Format',
    'txt': 'Text (encoded)',
    'html': 'XHTML Writer File',
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
            '--accept=socket,host=localhost,port=%d;urp;StarOffice.ServiceManager' % port,
            '--userid="%s"' % home_dir,
            '--norestore', '--nofirststartwizard', '--nologo', '--nocrashreport',
            '--nodefault', '--quickstart', '--norestart', '--nolockcheck', '--headless']
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

    context = uno.getComponentContext()
    svc_mgr = context.ServiceManager
    resolver = svc_mgr.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', context)
    connection_params = 'uno:socket,host=%s,port=%s;urp;' \
                        'StarOffice.ComponentContext' % ('localhost', port)

    uno_context = None
    for times in xrange(5):
        try:
            uno_context = resolver.resolve(connection_params)
            break
        except NoConnectException:
            time.sleep(1)

    if not uno_context:
        exit(1)

    uno_svc_mgr = uno_context.ServiceManager
    desktop = uno_svc_mgr.createInstanceWithContext('com.sun.star.frame.Desktop', uno_context)
    return popen, context, desktop


def get_free_port():
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.bind(('0.0.0.0', 0))
    _, port = s.getsockname()
    s.close()
    return port


def convert(source, target, target_format):
    home_dir = tempfile.mkdtemp()
    return_code = 1

    try:
        port = get_free_port()
        popen, context, desktop = start_openoffice(home_dir, port)
        
        input_stream = context.ServiceManager.createInstanceWithContext(
                'com.sun.star.io.SequenceInputStream', context)
        input_stream.initialize((uno.ByteSequence(source.read()),))

        doc = desktop.loadComponentFromURL('private:stream', '_blank', 0, to_properties({
           'InputStream': input_stream,
        }))
        doc.storeToURL('private:stream', to_properties({
            'FilterName': FILTER_MAP[target_format],
            'OutputStream': OutputStream(target)
        }))

        doc.dispose()
        doc.close(True)
        desktop.terminate()

        exit(popen.wait())
    except Exception as e:
        print >> sys.stderr, e
        exit(1)
    finally:
        shutil.rmtree(home_dir)


def main():
    parser = OptionParser()
    parser.add_option('-s', '--source', dest='source_file',
            help='Source file or `-` for stdin', metavar='FILE')
    parser.add_option('-t', '--target', dest='target_file',
            help='Target file or `-` for stdout', metavar='FILE')
    parser.add_option('-f', '--format', dest='target_format',
            help='Target format of the file', metavar='FORMAT')
    (options, args) = parser.parse_args()

    source_file = options.source_file
    target_file = options.target_file
    target_format = options.target_format

    if target_format is None:
        parser.error('Target format is required.')
    if source_file is None:
        parser.error('Source file is not specified.')
    if target_file is None:
        parser.error('Target file is not specified.')

    if source_file == '-':
        source = sys.stdin
    else:
        source = open(source_file, 'r')

    if target_file == '-':
        target = sys.stdout
    else:
        target = open(source_file, 'r')

    convert(source, target, target_format)


if __name__ == '__main__':
    main()
