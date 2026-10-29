function include(name) { return HtmlService.createHtmlOutputFromFile(name).getContent(); }
function doGet() { return HtmlService.createTemplateFromFile('Index').evaluate().setTitle('วนดี — บันทึกการขายรีไซเคิล').addMetaTag('viewport','width=device-width, initial-scale=1'); }
function setup() {
  const props=PropertiesService.getScriptProperties();
  if (!props.getProperty('SHEET_ID')) {
    const book=SpreadsheetApp.create('วนดี - รายการขายรีไซเคิล');
    props.setProperty('SHEET_ID',book.getId());
    book.getSheets()[0].setName('Sales').appendRow(['ID','เลขที่รายการ','วันที่ขาย','บริษัท','น้ำหนักรวม','ยอดขาย','รับเงินจริง','ข้อมูล JSON']);
  }
  if (!props.getProperty('FOLDER_ID')) props.setProperty('FOLDER_ID',DriveApp.createFolder('วนดี - หลักฐานการขาย').getId());
  return 'ตั้งค่าเรียบร้อย';
}
function sheet_(){const id=PropertiesService.getScriptProperties().getProperty('SHEET_ID');if(!id)throw new Error('กรุณารัน setup ก่อน');return SpreadsheetApp.openById(id).getSheetByName('Sales');}
function listSales(){const s=sheet_();return s.getLastRow()<2?[]:s.getRange(2,8,s.getLastRow()-1,1).getValues().map(r=>JSON.parse(r[0]));}
function saveSale(s){
  if(!s || !s.id || !s.date || !String(s.company||'').trim() || !String(s.author||'').trim() || !String(s.department||'').trim() || !Array.isArray(s.items)||!s.items.length)throw new Error('ข้อมูลการขายไม่ครบ');
  let weight=0,total=0;
  s.items.forEach(i=>{['quantity','weight','price'].forEach(k=>{if(i[k]===''||!isFinite(Number(i[k]))||Number(i[k])<0)throw new Error('จำนวน น้ำหนัก และราคาต้องไม่ติดลบ');});if(Number(i.weight)<=0)throw new Error('น้ำหนักต้องมากกว่า 0');weight+=Number(i.weight);total+=Math.round(Number(i.weight)*Number(i.price)*100)/100;});
  total=Math.round(total*100)/100;
  if(!isFinite(Number(s.received))||Number(s.received)<0||Number(s.received)>total)throw new Error('ยอดรับเงินไม่ถูกต้อง');
  if(Number(s.received)>0&&!s.paymentDate)throw new Error('กรุณาระบุวันที่รับเงิน');
  if(s.deliveryStatus==='ส่งมอบแล้ว'&&(!s.deliveryDate||!s.responsible))throw new Error('ข้อมูลส่งมอบไม่ครบ');
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const sheet=sheet_(),rows=sheet.getLastRow()<2?[]:sheet.getRange(2,1,sheet.getLastRow()-1,8).getValues(),index=rows.findIndex(r=>r[0]===s.id);
    const old=index<0?null:JSON.parse(rows[index][7]);
    if(old&&old.updatedAt>s.updatedAt)throw new Error('ข้อมูลถูกแก้ไขแล้ว กรุณาโหลดใหม่');
    const folder=DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('FOLDER_ID'));
    s.attachments=(s.attachments||[]).map(a=>{
      if(!a.data){const found=old&&(old.attachments||[]).find(f=>f.id===a.id);if(!found)throw new Error('ไฟล์แนบไม่ถูกต้อง');return found;}
      if(!/^image\//.test(a.type)&&a.type!=='application/pdf')throw new Error('ประเภทไฟล์ไม่รองรับ');
      const bytes=Utilities.base64Decode(a.data);if(bytes.length>10*1024*1024)throw new Error('ไฟล์เกิน 10 MB');
      const f=folder.createFile(Utilities.newBlob(bytes,a.type,a.name));return {id:f.getId(),name:f.getName(),size:bytes.length,url:f.getUrl()};
    });
    s.updatedAt=new Date().toISOString();const json=JSON.stringify(s);if(json.length>49000)throw new Error('ข้อมูลรายการยาวเกินขนาดเซลล์');
    const safe=v=>/^[=+@-]/.test(String(v))?"'"+v:v;
    const values=[s.id,safe(s.number),s.date,safe(s.company),weight,total,Number(s.received),json];
    if(index<0)sheet.appendRow(values);else sheet.getRange(index+2,1,1,8).setValues([values]);return {ok:true};
  }finally{lock.releaseLock();}
}
