#!/usr/bin/env python3
"""Integration check: intercepted API, real browser downloads, XLSM/format inspection.
Requires playwright, openpyxl, Chromium. Does not execute VBA or contact Google.
"""
import asyncio, json, zipfile
from urllib.parse import urlsplit
from pathlib import Path
from xml.etree import ElementTree as ET
from playwright.async_api import async_playwright, expect
import openpyxl
ROOT=Path(__file__).resolve().parent.parent
import os, tempfile
OUT=Path(os.environ.get('GTO_EXPORT_TEST_OUTPUT', tempfile.mkdtemp(prefix='gto-export-test-')))
OUT.mkdir(parents=True, exist_ok=True)
DISC=[dict(col='F',name='прыжки'),dict(col='G',name='длинная')]
NORMS=[dict(key='V|мужской|прыжки',name='Прыжок в длину с места',dir='больше',fmt='число',b=180,s=200,g=220),dict(key='V|мужской|длинная',name='Бег на 2000 м',dir='меньше',fmt='мин:сек',b=900,s=850,g=800)]
DATA=[dict(row=i+2,num=i+1,fio='Проверка Участник '+str(i+1),uin='0012345678'+str(i),stupen='V ступень (14-15 лет)',pol='мужской',klass='9',letter='А',present=True,res=res) for i,res in enumerate([['230','12:40'],['190','14,20'],['0','28:05'],['','']])]
async def main():
 async with async_playwright() as p:
  b=await p.chromium.launch(executable_path='/usr/bin/chromium',headless=True,args=['--no-sandbox'])
  page=await b.new_page(accept_downloads=True);errors=[]
  page.on('pageerror',lambda e:errors.append(str(e)))
  async def route(r):
   if r.request.url.startswith('http://gto.test/'):
    name=urlsplit(r.request.url).path.lstrip('/') or 'results.html'
    path=ROOT/name
    return await r.fulfill(body=path.read_bytes(),content_type='text/html' if name.endswith('.html') else 'text/css' if name.endswith('.css') else 'text/javascript')
   if 'script.google.com' in r.request.url: return await r.fulfill(json=dict(ok=True,data=DATA,disciplines=DISC,norms=NORMS,event='01.11.2026'),headers={'Access-Control-Allow-Origin':'*'})
   return await r.abort()
  await page.route('**/*',route)
  await page.goto('http://gto.test/results.html')
  await expect(page.locator('#sTotal')).to_have_text('4')
  for button,name in [('#xlsBtn','results'),('#blankBtn','blank')]:
   async with page.expect_download(timeout=60000) as d:
    await page.locator(button).click()
   download=await d.value
   suffix='.xlsm' if name=='results' else '.xlsx'
   assert download.suggested_filename.endswith(suffix)
   await download.save_as(OUT/(name+suffix))
   check(OUT/(name+suffix))
  assert not errors,errors
  await b.close()
def check(path):
 w=openpyxl.load_workbook(path,keep_vba=path.suffix=='.xlsm')
 count=0
 for s in w:
  for dim in s.column_dimensions.values():assert dim.number_format=='General',(s.title,dim.index,dim.number_format)
  for dim in s.row_dimensions.values():assert dim.number_format=='General',(s.title,dim.index,dim.number_format)
  for row in s:
   for c in row:
    assert c.number_format=='General',(s.title,c.coordinate,c.number_format)
    count+=1
 ns={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
 with zipfile.ZipFile(path) as z:
  styles=ET.fromstring(z.read('xl/styles.xml'))
  assert not styles.findall('m:numFmts/m:numFmt',ns),'custom formats remain'
  for xf in styles.findall('m:cellXfs/m:xf',ns):assert xf.get('numFmtId','0')=='0',xf.attrib
  for fmt in styles.findall('m:dxfs/m:dxf/m:numFmt',ns):assert fmt.get('numFmtId','0')=='0',fmt.attrib
 if path.stem=='results':
  with zipfile.ZipFile(path) as z:
   assert z.read('xl/vbaProject.bin')==(ROOT/'vendor/protocol/vbaProject.bin').read_bytes()
   assert b'macroEnabled.main+xml' in z.read('[Content_Types].xml')
   assert b'relationships/vbaProject' in z.read('xl/_rels/workbook.xml.rels')
  assert 'GTO_ProtocolData' in w.defined_names
  assert sum(not w['Протокол'].row_dimensions[r].hidden for r in range(10,513))==3
  with zipfile.ZipFile(path) as z:
   assert b'codeName="ThisWorkbook"' in z.read('xl/workbook.xml')

  assert w['Результаты']['D2'].value=='00123456780'
  assert [w['Результаты'][f'I{r}'].value for r in range(2,5)]==['12:40','14:20','28:05']
  assert [w['Нормативы'][f'{c}3'].value for c in 'GHI']==['15:00','14:10','13:20']
 print('PASS',path.name,len(w.sheetnames),'sheets;',count,'cells General; column/row styles and conditional formats General; time and UIN preserved')
if __name__ == '__main__':
 asyncio.run(main())
