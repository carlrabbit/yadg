import hashlib
import json
import os
import platform
import subprocess
import time
import tempfile
import uno
from com.sun.star.awt import Size, Point
from com.sun.star.text import ControlCharacter
from com.sun.star.beans import PropertyValue

repo = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
fixture_dir = os.path.join(repo, 'tests', 'fixtures', 'm0010')
os.makedirs(fixture_dir, exist_ok=True)
path = os.path.join(fixture_dir, 'libreoffice-origin.docx')
soffice = r'C:\Program Files\LibreOffice\program\soffice.exe'
profile = tempfile.mkdtemp(prefix='yadg-m0010-lo-profile-')
proc = subprocess.Popen([soffice, '--headless', '--nologo', '--nodefault', '--nofirststartwizard', '--accept=socket,host=127.0.0.1,port=2002;urp;StarOffice.ComponentContext', '-env:UserInstallation=file:///' + profile.replace('\\', '/')])
version = '26.8.0.3'
try:
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
    ctx = None
    for _ in range(60):
        try:
            ctx = resolver.resolve('uno:socket,host=127.0.0.1,port=2002;urp;StarOffice.ComponentContext')
            break
        except Exception:
            time.sleep(0.25)
    if ctx is None:
        raise RuntimeError('Unable to connect to LibreOffice Writer UNO runtime')
    desktop = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)
    doc = desktop.loadComponentFromURL('private:factory/swriter', '_blank', 0, ())
    print('loaded', flush=True)
    paragraph_styles = doc.getStyleFamilies().getByName('ParagraphStyles')
    for style_name in ['Heading1', 'Heading2', 'Heading3', 'Heading4', 'Heading5', 'Heading6', 'Heading7', 'Heading8', 'Heading9']:
        if not paragraph_styles.hasByName(style_name):
            style = doc.createInstance('com.sun.star.style.ParagraphStyle')
            paragraph_styles.insertByName(style_name, style)
        else:
            style = paragraph_styles.getByName(style_name)
        level = int(style_name[-1])
        style.OutlineLevel = level
        style.CharHeight = max(11.0, 20.0 - level)
        style.CharWeight = 150.0
    text = doc.getText()
    cursor = text.createTextCursor()
    def insert_lines(lines):
        for index, line in enumerate(lines):
            text.insertString(cursor, line, False)
            if index < len(lines) - 1:
                text.insertControlCharacter(cursor, ControlCharacter.PARAGRAPH_BREAK, False)
    insert_lines(['{{yadg:frontmatter}}', 'version: 1', 'styles:', '  headings:', '    1: Heading1', '    2: Heading2', '    3: Heading3', '    4: Heading4', '    5: Heading5', '    6: Heading6', '    7: Heading7', '    8: Heading8', '    9: Heading9', '{{/yadg:frontmatter}}'])
    text.insertControlCharacter(cursor, ControlCharacter.PARAGRAPH_BREAK, False)
    insert_lines(['LibreOffice-origin static cover content.', 'Template-owned section', '{{section:architecture}}', 'Template-owned paragraph between placements.', 'Second template-owned section', '{{content:architecture}}', 'LibreOffice-origin static closing content.'])
    print('text', flush=True)
    paragraphs = text.createEnumeration()
    while paragraphs.hasMoreElements():
        p = paragraphs.nextElement()
        if hasattr(p, 'String') and p.String in ('Template-owned section', 'Second template-owned section'):
            p.ParaStyleName = 'Heading4'

    table = doc.createInstance('com.sun.star.text.TextTable')
    table.initialize(2, 2)
    text.insertTextContent(cursor, table, False)
    table.getCellByName('A1').String = 'Template table'
    table.getCellByName('B1').String = '{{value:story-value}}'
    table.getCellByName('A2').String = 'Preserved'
    table.getCellByName('B2').String = 'Layout'
    print('table', flush=True)
    text.insertControlCharacter(cursor, ControlCharacter.PARAGRAPH_BREAK, False)

    footnote = doc.createInstance('com.sun.star.text.Footnote')
    text.insertTextContent(cursor, footnote, False)
    footnote.String = 'Footnote {{value:story-value}}'
    endnote = doc.createInstance('com.sun.star.text.Endnote')
    text.insertTextContent(cursor, endnote, False)
    endnote.String = 'Endnote {{value:story-value}}'
    annotation = doc.createInstance('com.sun.star.text.textfield.Annotation')
    annotation.Content = 'Comment {{value:story-value}}'
    annotation.Author = 'YADG synthetic fixture'
    text.insertTextContent(cursor, annotation, False)

    frame = doc.createInstance('com.sun.star.text.TextFrame')
    frame.setSize(Size(6000, 1800))
    frame.setPropertyValue('AnchorType', 0)
    text.insertTextContent(cursor, frame, False)
    frame.Text.String = 'Text box {{value:story-value}}'
    print('stories', flush=True)

    styles = doc.getStyleFamilies().getByName('PageStyles')
    names = styles.getElementNames()
    default_name = 'Default Page Style' if styles.hasByName('Default Page Style') else names[0]
    default = styles.getByName(default_name)
    default.HeaderIsOn = True
    default.FooterIsOn = True
    default.HeaderText.String = 'Static header {{value:story-value}}'
    default.FooterText.String = 'Static footer {{value:story-value}}'
    try:
        first = styles.getByName('First Page')
        first.HeaderIsOn = True
        first.FooterIsOn = True
        first.HeaderText.String = 'First-page {{value:story-value}}'
        first.FooterText.String = 'First-footer {{value:story-value}}'
    except Exception:
        pass
    print('styles', flush=True)

    props = (PropertyValue('FilterName', 0, 'Office Open XML Text', 0), PropertyValue('Overwrite', 0, True, 0))
    doc.storeAsURL(uno.systemPathToFileUrl(path), props)
    print('stored', flush=True)
    doc.close(True)
    digest = hashlib.sha256(open(path, 'rb').read()).hexdigest().upper()
    record = {'path': 'tests/fixtures/m0010/libreoffice-origin.docx', 'application': 'LibreOffice Writer', 'version': version, 'platform': platform.platform(), 'method': 'Created as a new document through LibreOffice Writer UNO document model and saved by Writer as DOCX', 'date': time.strftime('%Y-%m-%dT%H:%M:%S%z'), 'sha256': digest, 'synthetic': True, 'nonConfidential': True}
    with open(os.path.join(fixture_dir, 'libreoffice-origin.provenance.json'), 'w', encoding='utf-8') as f:
        json.dump(record, f, indent=2)
finally:
    subprocess.run(['taskkill', '/PID', str(proc.pid), '/T', '/F'], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        proc.wait(timeout=10)
    except subprocess.TimeoutExpired:
        proc.kill()
