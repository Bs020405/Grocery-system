// Billing page logic (P and PRE are set in billing.html)
function addRow(){
  const tr = document.createElement('tr');
  let opts = '<option value="">-- select --</option>';
  P.forEach((p,i)=>{opts += `<option value="${p.id}" data-i="${i}">${p.name} (&#8377;${p.price}, stock ${p.stock})</option>`});
  tr.innerHTML = `<td><select name="product_id" onchange="calc()">${opts}</select></td>
   <td><input name="qty" type="number" step="0.01" min="0" value="1" oninput="calc()" style="width:90px"></td>
   <td class="right amt">0.00</td>`;
  document.getElementById('rows').appendChild(tr);
}
function calc(){
  let sub = 0, gst = 0;
  document.querySelectorAll('#rows tr').forEach(tr=>{
    const sel = tr.querySelector('select'), q = parseFloat(tr.querySelector('input').value)||0;
    let line = 0;
    if(sel.value){ const p = P[sel.selectedOptions[0].dataset.i]; line = p.price*q; gst += line*p.gst/100; }
    sub += line; tr.querySelector('.amt').textContent = line.toFixed(2);
  });
  const d = parseFloat(document.getElementById('disc').value)||0;
  document.getElementById('sub').textContent = sub.toFixed(2);
  document.getElementById('gst').textContent = gst.toFixed(2);
  document.getElementById('total').textContent = Math.max(0, sub+gst-d).toFixed(2);
}
addRow();
if (PRE) { const s = document.querySelector('#rows select'); s.value = String(PRE); calc(); }
