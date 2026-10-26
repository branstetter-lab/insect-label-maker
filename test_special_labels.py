"""Integration checks for compact identifiers and determination labels."""
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
import openpyxl
from pypdf import PdfReader
import zxingcpp

class SpecialLabels(unittest.TestCase):
    def run_mode(self, root, headers, rows, mode, *extra):
        source=Path(root)/'input.xlsx';out=Path(root)/'labels.pdf'
        book=openpyxl.Workbook();sheet=book.active;sheet.append(headers)
        for row in rows:sheet.append(row)
        book.save(source)
        result=subprocess.run([sys.executable,'label_maker.py',str(source),'--mode',mode,'--output',str(out),*extra],capture_output=True,text=True)
        return result,out
    def test_generic_ids_only_and_omit_text(self):
        with tempfile.TemporaryDirectory() as root:
            result,out=self.run_mode(root,['Catalog_Number','Field_Number'],[['CASENT1234567','MUST-NOT-PRINT']],'barcode')
            self.assertEqual(result.returncode,0,result.stderr)
            report=json.loads(out.with_suffix('.json').read_text());self.assertEqual(report['width_inches'],.55);self.assertEqual(report['height_inches'],.30)
            page=PdfReader(out).pages[0];text=page.extract_text()
            self.assertIn('CASENT1234567',text);self.assertNotIn('MUST-NOT-PRINT',text)
            self.assertIn('Identifier',text);self.assertFalse(report['field_headings'])
            images=list(page.images);self.assertEqual(len(images),1)
            self.assertEqual(zxingcpp.read_barcode(images[0].image).text,'CASENT1234567')
            result,out=self.run_mode(root,['Catalog_Number'],[['CASENT1234567']],'barcode','--omit-catalog-text')
            self.assertEqual(result.returncode,0,result.stderr)
            self.assertNotIn('CASENT1234567',PdfReader(out).pages[0].extract_text())
    def test_determination_required_fields_and_author(self):
        with tempfile.TemporaryDirectory() as root:
            result,out=self.run_mode(root,['Catalog_Number','Genus','Species','Author','Determiner_Label'],[['DEMO000001','Lasius','niger','(Linnaeus, 1758)','A. Smith 2026']],'determination')
            self.assertEqual(result.returncode,0,result.stderr)
            report=json.loads(out.with_suffix('.json').read_text());lines=report['records'][0]['lines']
            self.assertEqual(lines[0]['italic_end'],12)  # Genus and species, not authority.
            self.assertEqual(lines[0]['text'], 'Lasius niger')
            self.assertEqual(lines[1]['text'], '(Linnaeus, 1758)')
            self.assertEqual(lines[1]['italic_end'], 0)
            self.assertIn('Det. A. Smith 2026',PdfReader(out).pages[0].extract_text())
            result,out=self.run_mode(root,['Catalog_Number','Genus','Species','Author','Determiner_Label'],[['DEMO000001','Lasius','niger','L.','Det. A. Smith 2026']],'determination')
            self.assertEqual(result.returncode,0,result.stderr)
            lines=json.loads(out.with_suffix('.json').read_text())['records'][0]['lines']
            self.assertEqual(lines[0]['text'], 'Lasius niger L.')
            self.assertNotIn('Det. Det.', PdfReader(out).pages[0].extract_text())
            result,out=self.run_mode(root,['Catalog_Number','Genus','Species'],[['DEMO000001','Lasius','niger']],'determination')
            self.assertNotEqual(result.returncode,0);self.assertIn('Determiner_Label is required',result.stderr)
    def test_unknown_taxon_qualifier_stays_upright_and_fit_errors(self):
        with tempfile.TemporaryDirectory() as root:
            result,out=self.run_mode(root,['Catalog_Number','Genus','Species','Determiner_Label'],[['DEMO000001','Formica','sp.','A. Smith 2026']],'determination')
            self.assertEqual(result.returncode,0,result.stderr)
            self.assertEqual(json.loads(out.with_suffix('.json').read_text())['records'][0]['lines'][0]['italic_end'],7)
            result,out=self.run_mode(root,['Catalog_Number'],[['THIS-CODE-IS-TOO-LONG-TO-PRINT']],'barcode')
            self.assertNotEqual(result.returncode,0);self.assertIn('increase --height',result.stderr)
    def test_pin_barcode_has_no_caption_and_preserves_main_matrix_size(self):
        with tempfile.TemporaryDirectory() as root:
            result,out=self.run_mode(root,['Catalog_Number'],[['CASENT1234567']],'pin-barcode')
            self.assertEqual(result.returncode,0,result.stderr)
            d=json.loads(out.with_suffix('.json').read_text())
            self.assertEqual(d['barcode_scale'],1.02)
            self.assertEqual(d['barcode_rendering'],'raster')
            self.assertEqual(d['barcode_raster_dpi'],1200)
            self.assertEqual(d['records'][0]['lines'],[])
            self.assertGreaterEqual(d['pin_blank_width_inches'],.08)
            self.assertEqual(PdfReader(out).pages[0].extract_text().strip(),'CASENT1234567')
            result,out=self.run_mode(root,['Catalog_Number'],[['CASENT1234567']],'pin-barcode','--width','.25')
            self.assertNotEqual(result.returncode,0)
            self.assertIn('pin space',result.stderr)
    def test_optional_method_and_three_line_limit(self):
        import label_maker as l
        l.pdfmetrics.registerFont(l.TTFont('LabelRegular', l.FONT_PATH))
        l.pdfmetrics.registerFont(l.TTFont('LabelItalic', l.ITALIC_FONT_PATH))
        row={'Catalog_Number':'DEMO000001','_excel_row':2,'MacroHabitat_Label':'meadow','Microhabitat_Label':'Phacelia sp.','Sampling_Protocol':'net'}
        p=l.habitat_label_plan(row,.62*72,.20*72,include_method=True)
        self.assertEqual(p['lines'],['meadow','Phacelia sp.','net'])
        self.assertEqual(p['line_sources'],['macrohabitat','microhabitat','method'])
        self.assertEqual(p['italic_ranges'][-1],[])
        self.assertEqual(len(l.habitat_label_plan(row,.62*72,.20*72)['lines']),2)
        row['Sampling_Protocol']='net sweep repeated over a very long meadow path'
        with self.assertRaisesRegex(ValueError,'maximum is 3'):
            l.habitat_label_plan(row,.62*72,.20*72,include_method=True)
    def test_vector_barcode_option_is_rejected(self):
        with tempfile.TemporaryDirectory() as root:
            result,_=self.run_mode(root,['Catalog_Number'],[['DEMO000001']],'barcode','--barcode-rendering','vector')
            self.assertNotEqual(result.returncode,0)
            self.assertIn('invalid choice',result.stderr)
if __name__=='__main__':unittest.main()
