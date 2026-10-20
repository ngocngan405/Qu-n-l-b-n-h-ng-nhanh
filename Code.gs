/** Google Apps Script: tạo các sheet Products, Customers, Orders, OrderItems và lưu snapshot từ ứng dụng. */
function doPost(e) {
  try {
    var data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (data.action !== 'replaceAll') throw new Error('Unsupported action');
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    writeObjects_(ss, 'Products', data.products || [], ['id','name','sku','category','unit','price','cost','stock']);
    writeObjects_(ss, 'Customers', data.customers || [], ['id','name','phone','points']);
    writeObjects_(ss, 'Orders', (data.orders || []).map(function(o){return {id:o.id,date:o.date,customer:o.customer,note:o.note,total:o.total,status:o.status,updatedAt:data.updatedAt};}), ['id','date','customer','note','total','status','updatedAt']);
    var lines=[];
    (data.orders||[]).forEach(function(o){(o.items||[]).forEach(function(i){lines.push({orderId:o.id,date:o.date,customer:o.customer,productId:i.id,name:i.name,sku:i.sku||'',unit:i.unit||'',qty:i.qty,price:i.price,subtotal:i.subtotal});});});
    writeObjects_(ss, 'OrderItems', lines, ['orderId','date','customer','productId','name','sku','unit','qty','price','subtotal']);
    var settings=ss.getSheetByName('Settings')||ss.insertSheet('Settings');settings.clearContents();settings.getRange(1,1,4,2).setValues([['key','value'],['shopName',data.shopName||'Cửa hàng của tôi'],['updatedAt',data.updatedAt||new Date().toISOString()],['note','Được cập nhật từ ứng dụng quản lý bán hàng']]);
    return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
  } catch(err) { return ContentService.createTextOutput(JSON.stringify({ok:false,error:String(err)})).setMimeType(ContentService.MimeType.JSON); }
}
function writeObjects_(ss, name, objects, keys) {
  var sh=ss.getSheetByName(name)||ss.insertSheet(name); sh.clearContents();
  sh.getRange(1,1,1,keys.length).setValues([keys]);
  if(objects.length) sh.getRange(2,1,objects.length,keys.length).setValues(objects.map(function(o){return keys.map(function(k){var v=o[k];return v===undefined||v===null?'':v;});}));
  sh.setFrozenRows(1); sh.getRange(1,1,1,keys.length).setFontWeight('bold');
}
