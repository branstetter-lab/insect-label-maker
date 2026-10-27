"""Compose real collection and fictional determination example PDFs at actual size onto one comparison sheet."""
from pathlib import Path
import copy,io,json
from reportlab.pdfgen import canvas
from pypdf import PdfReader,PdfWriter,Transformation
ROOT=Path(__file__).resolve().parents[1]
OPTIONS=[
('Locality: separate lines','sampler_real_lines','--mode locality'),
('Locality: wrapped text','sampler_real_flow','--layout flow'),
('Two labels: habitat','sampler_real_paired','--separate-habitat'),
('Two labels: habitat + method','sampler_real_method','--separate-habitat --include-method'),
('Generic barcode','sampler_real_barcode','--mode barcode'),
('Determination + barcode (fictional)','example_labels_determinations','--mode determination'),
('Pin barcode: blank space at left','sampler_real_pin','--mode pin-barcode'),
('Collection only: no barcode','example_labels_collection_only','--mode collection-only'),
('Collection only + habitat','example_labels_collection_only_paired','--mode collection-only --separate-habitat --include-method'),
('Collection only: wrapped','example_labels_collection_only_flow','--mode collection-only --layout flow')]
def main():
 stream=io.BytesIO();c=canvas.Canvas(stream,pagesize=(612,792));c.setTitle('Pittacia Insectorum — label options')
 c.setFont('Helvetica-Bold',16);c.drawString(36,756,'Pittacia Insectorum — label options')
 c.setFont('Helvetica',9);c.drawString(36,738,'Real collection example: BBSL1207977 / MGB4209. Shared with permission.')
 c.drawString(36,725,'Determination example is fictional. Repeated IDs are intentional. Print at Actual Size / 100%.')
 c.line(36,715,108,715);c.drawString(116,712,'1 inch / 25.4 mm calibration line')
 for i,(name,stem,command) in enumerate(OPTIONS):
  y=670-i*56
  c.setFont('Helvetica-Bold',10);c.drawString(36,y,name)
  c.setFont('Helvetica',8);c.drawString(36,y-15,command)
  d=json.loads((ROOT/'examples'/f'{stem}.json').read_text())
  c.drawString(36,y-28,f"Main: {d['width_inches']:g} × {d['height_inches']:g} in")
  if d.get('macrohabitat_omitted_count'):
   c.drawString(36,y-40,'Fit fallback: macrohabitat omitted.')
  if d.get('uncertainty_omitted_count'):
   c.drawString(36,y-40,'Fit fallback: coordinate uncertainty omitted.')
  top=680-i*56; w=d['width_inches']*72; h=d['height_inches']*72
  c.setStrokeColorRGB(.7,.7,.7); c.setLineWidth(.15);c.rect(420,top-h,w,h)
  if d.get('separate_habitat'):
   c.rect(420,top-h-d['habitat_label_gap_inches']*72-d['habitat_label_height_inches']*72,d['habitat_label_width_inches']*72,d['habitat_label_height_inches']*72)
  c.setStrokeColorRGB(0,0,0)
 c.setFont('Helvetica',8);c.drawString(36,65,'Workbook: examples/label_options_examples.xlsx — one worksheet per option; Instructions lists flags.')
 c.drawString(36,52,'Blank workbook: templates/all_options_input_template.xlsx. Keep pin holes outside barcode margins.')
 c.save();stream.seek(0);page=PdfReader(stream).pages[0]
 for i,(_,stem,_) in enumerate(OPTIONS):
  d=json.loads((ROOT/'examples'/f'{stem}.json').read_text());r=d['records'][0];pos=r['placement'];x=pos['x_points'];y=pos['y_points'];w=d['width_inches']*72;h=d['height_inches']*72
  habitat=r.get('habitat_label');bottom=habitat['y_points'] if habitat else y
  sample=copy.copy(PdfReader(ROOT/'examples'/f'{stem}.pdf').pages[0])
  sample.cropbox.lower_left=(x,bottom);sample.cropbox.upper_right=(x+w,y+h)
  sample.trimbox.lower_left=(x,bottom);sample.trimbox.upper_right=(x+w,y+h)
  target_top=680-i*56
  page.merge_transformed_page(sample,Transformation().translate(tx=420-x,ty=target_top-(y+h)))
 writer=PdfWriter();writer.add_page(page);writer.add_metadata({'/Title':'Pittacia Insectorum — label options','/Author':'Pittacia Insectorum'})
 output=ROOT/'examples/label_options_sampler.pdf';writer.write(output);print(output)
if __name__=='__main__':main()
