import html
import sys
import os
import shutil
import tempfile
import subprocess
import clipboard
from docx import Document
from lxml import etree


MATH_NAMESPACE = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
WORD_NAMESPACE = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
ONENOTE_MATH_NAMESPACE = 'http://schemas.microsoft.com/office/2004/12/omml'


def _prepare_markdown(latex_string):
    text = latex_string.strip()
    has_math_delimiters = (
        '$' in text or r'\(' in text or r'\[' in text
    )
    if has_math_delimiters:
        return text
    return f'${text}$'


def _paragraph_to_html(paragraph):
    has_text = any(
        (element.text or '').replace('\u200b', '').strip()
        for element in paragraph._p.iter(f'{{{WORD_NAMESPACE}}}t')
    )

    def render(element):
        if element.tag == f'{{{MATH_NAMESPACE}}}oMath':
            raw_omml = etree.tostring(element, encoding='unicode').replace(
                MATH_NAMESPACE, ONENOTE_MATH_NAMESPACE
            )
            if has_text:
                return f'<!--[if gte msEquation 12]>{raw_omml}<![endif]-->'

            omml_para = (
                f'<m:oMathPara xmlns:m="{ONENOTE_MATH_NAMESPACE}">'
                f'{raw_omml}</m:oMathPara>'
            )
            return f'<!--[if gte msEquation 12]>{omml_para}<![endif]-->'

        if element.tag == f'{{{WORD_NAMESPACE}}}t':
            return html.escape(element.text or '', quote=False)
        if element.tag in (f'{{{WORD_NAMESPACE}}}br', f'{{{WORD_NAMESPACE}}}cr'):
            return '<br/>'
        if element.tag == f'{{{WORD_NAMESPACE}}}tab':
            return '&#9;'
        return ''.join(render(child) for child in element)

    return f'<p>{"".join(render(child) for child in paragraph._p)}</p>'


def latex_to_onenote_clipboard(latex_string):
    """
    Convert LaTeX equations to OMML and place the resulting text on the clipboard
    """
    markdown = _prepare_markdown(latex_string)
    
    # Create temp dir and md file for pandoc conversion
    with tempfile.TemporaryDirectory() as tmpdir:
        md_file = os.path.join(tmpdir, 'input.md')
        docx_file = os.path.join(tmpdir, 'output.docx')
        
        with open(md_file, 'w', encoding='utf-8') as f:
            f.write(markdown)
        
        # Convert with pandoc to docx (which contains OMML)
        _local_pandoc = os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Pandoc', 'pandoc.exe')
        pandoc_path = (
            shutil.which('pandoc') or
            (_local_pandoc if os.path.exists(_local_pandoc) else None) or
            'pandoc'
        )

        try:
            result = subprocess.run(
                [pandoc_path, md_file, '-o', docx_file],
                capture_output=True,
                text=True,
                check=True
            )
        except subprocess.CalledProcessError as e:
            print(f"Pandoc error: {e.stderr}")
            return False
        except FileNotFoundError:
            print("Error: Pandoc not found. Please install Pandoc first.")
            return False
        
        # Open docx and extract OMML
        try:
            doc = Document(docx_file)
            
            html_output = ''.join(
                _paragraph_to_html(paragraph) for paragraph in doc.paragraphs
            )
            if '<!--[if gte msEquation 12]>' not in html_output:
                print("Warning: No equation found in converted document")
                return False

            clipboard.set_clipboard(
                html_output,
                format=clipboard.ClipboardFormat.HTML_Format
            )
            print("✓ Converted LaTeX to OMML and copied to clipboard!")
            print(f"Original input: {latex_string.strip()}")
            print("Now paste into OneNote with Ctrl+V")
            return True

        except Exception as e:
            print(f"Error extracting OMML: {e}")
            return False



if __name__ == "__main__":
    # Handle arguments
    force_convert = '-f' in sys.argv
    args = [arg for arg in sys.argv[1:] if arg != '-f']

    # Get LaTeX from clipboard or command line argument
    if len(args) > 0:
        latex_input = args[0]
    else:
        try:
             # Try CF_UNICODETEXT
             latex_input = clipboard.get_clipboard(clipboard.ClipboardFormat.CF_UNICODETEXT)
        except:
             # Fallback to CF_TEXT
             latex_input = clipboard.get_clipboard(clipboard.ClipboardFormat.CF_TEXT)
        
        print(f"Reading from clipboard: {latex_input}")
    
    if not latex_input or not latex_input.strip():
        print("Error: No LaTeX input provided")
        print("Usage: python latex_to_onenote.py 'V_{rms}=\\frac{V_{peak}}{\\sqrt{2}}'")
        print("   or: Copy LaTeX to clipboard and run: python latex_to_onenote.py")
        print("   (Use -f to force conversion even if no LaTeX detected)")
        sys.exit(1)

    # Basic LaTeX detection (unless forced)
    common_latex_symbols = ['\\', '{', '}', '_', '^', '$']
    is_latex = any(symbol in latex_input for symbol in common_latex_symbols)
    
    if not is_latex and not force_convert:
        print("\nError: No LaTeX symbols detected in input.")
        print("The input doesn't look like LaTeX (missing '\\', '{', '}', '_', '^', or '$').")
        print("Use -f to force conversion anyway.")
        sys.exit(1)
    
    success = latex_to_onenote_clipboard(latex_input)
    sys.exit(0 if success else 1)
