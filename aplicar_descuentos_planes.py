#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
aplicar_descuentos_planes.py  ·  PROCSIS: módulo de descuentos al crear/editar planes, con precio final calculado al instante.

Uso (en la carpeta de tu app.py):
    python aplicar_descuentos_planes.py            # parchea ./app.py
    python aplicar_descuentos_planes.py ruta/app.py

Qué hace:
  · En «Crear plan nuevo» y «Editar plan» aparece la tarjeta «Descuento y precio final»:
      - escribes el % de descuento, o el descuento en pesos, o el precio final que quieres cobrar, y los otros dos se calculan solos;
      - muestra el precio lista, el descuento, el PRECIO FINAL MENSUAL, el ahorro por mes y el ahorro durante la vigencia;
      - al guardar, el sistema ya guarda ese precio final (el servidor lo calcula igual: lista × (1 − %)).
  · No cambia tu base de datos ni tus rutas: usa los campos que ya existen (precio, descuento_pct, descuento_meses).
Es idempotente, crea app.py.bak_descuentos_planes y no escribe nada si el resultado no compila.
"""
import shutil
import sys
import time
from pathlib import Path

MARCA = "[procsis-descuentos-planes]"

BLOQUE = r'''
# [procsis-descuentos-planes]
# ═════════════ PLANES: descuento + precio final calculado al instante (crear y editar plan) ═════════════
_PFD_HTML = r"""
<style>
.pfd{margin:18px 0 6px;border:1px solid #bfdbfe;background:#f0f7ff;border-radius:14px;padding:16px}
.pfd-t{font-weight:800;color:#0020C2;font-size:14px;margin-bottom:4px}
.pfd-s{font-size:12px;color:#64748b;margin:0 0 8px}
.pfd-res{margin-top:12px;background:#fff;border:1px solid #dbeafe;border-radius:12px;padding:12px 14px;font-size:13px;color:#1e293b;line-height:1.7}
.pfd-final{font-size:22px;font-weight:800;color:#047857}
.pfd-tag{display:inline-block;background:#dcfce7;color:#166534;border-radius:999px;padding:1px 9px;font-size:11px;font-weight:700;margin-left:6px}
.pfd-warn{color:#b45309;font-weight:700}
</style>
<div class="pfd" id="pfd">
  <div class="pfd-t">💸 Descuento y precio final</div>
  <p class="pfd-s">Escribe <b>una</b> de las tres opciones y las otras se calculan solas. El plan se guarda con el precio final ya descontado.</p>
  <div class="row2">
    <div><label>Descuento en pesos (COP por mes)</label><input id="pfd-valor" type="number" min="0" step="any" placeholder="Ej.: 50000"></div>
    <div><label>Precio final que quiero cobrar (COP por mes)</label><input id="pfd-final" type="number" min="0" step="any" placeholder="Ej.: 199000"></div>
  </div>
  <div class="pfd-res" id="pfd-res">Escribe el precio lista para ver el resultado.</div>
</div>
<script>
(function(){
 var q=function(n){return document.querySelector('[name="'+n+'"]');};
 var L=q('precio'),P=q('descuento_pct'),M=q('descuento_meses'),V=document.getElementById('pfd-valor'),F=document.getElementById('pfd-final'),R=document.getElementById('pfd-res');
 if(!L||!P||!V||!F||!R)return;
 P.step='any';
 var cop=function(n){return '$'+Math.round(n).toLocaleString('es-CO');};
 var num=function(x){var n=parseFloat(x);return isNaN(n)?0:n;};
 var r4=function(n){return Math.round(n*10000)/10000;};
 function pintar(){
  var lista=num(L.value),pct=Math.max(0,Math.min(100,num(P.value))),meses=Math.max(0,Math.round(num(M&&M.value)));
  if(lista<=0){R.textContent='Escribe el precio lista para ver el resultado.';return;}
  var desc=lista*pct/100,fin=lista-desc,h='';
  h+='Precio lista: <b>'+cop(lista)+'</b> al mes<br>';
  if(pct>0){
   h+='Descuento: <b>'+r4(pct)+'%</b> (−'+cop(desc)+' al mes)<br>';
   h+='Precio final con descuento: <span class="pfd-final">'+cop(fin)+'</span><span class="pfd-tag">al mes</span><br>';
   if(meses>0){h+='Vigente durante <b>'+meses+' mes(es)</b>: ahorras '+cop(desc*meses)+'. Desde el mes '+(meses+1)+' se cobra '+cop(lista)+'.';}
   else{h+='Sin duración definida: el descuento queda vigente hasta que lo quites.';}
  }else{h+='Sin descuento: el precio final es <span class="pfd-final">'+cop(lista)+'</span><span class="pfd-tag">al mes</span>';}
  if(pct>=100)h+='<br><span class="pfd-warn">Con 100 % el plan queda gratis.</span>';
  R.innerHTML=h;
 }
 function desde(src){
  var lista=num(L.value);
  if(lista>0){
   if(src==='valor'&&V.value!==''){P.value=r4(Math.max(0,Math.min(100,num(V.value)/lista*100)));}
   if(src==='final'&&F.value!==''){P.value=r4(Math.max(0,Math.min(100,(1-num(F.value)/lista)*100)));}
  }
  var pct=Math.max(0,Math.min(100,num(P.value))),d=lista*pct/100;
  if(src!=='valor')V.value=(pct>0&&lista>0)?Math.round(d*100)/100:'';
  if(src!=='final')F.value=(pct>0&&lista>0)?Math.round((lista-d)*100)/100:'';
  pintar();
 }
 L.addEventListener('input',function(){desde('pct');});
 P.addEventListener('input',function(){desde('pct');});
 V.addEventListener('input',function(){desde('valor');});
 F.addEventListener('input',function(){desde('final');});
 if(M)M.addEventListener('input',pintar);
 desde('pct');
})();
</script>
"""

if "_plan_form_fields" in globals():
    _pfd_plan_form_fields_orig = _plan_form_fields

    def _plan_form_fields(p=None, crear=False):
        """Campos del plan + tarjeta de descuento con precio final al instante."""
        return _pfd_plan_form_fields_orig(p, crear) + _PFD_HTML


'''

ANCLAS = ["def _plan_form_fields(", "def _aplicar_campos_plan", 'name="descuento_pct"', 'name="precio"', 'name="descuento_meses"']


def main():
    ruta = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("app.py")
    if not ruta.is_file():
        sys.exit("No encuentro %s. Ejecuta el script en la carpeta de tu app.py o pásale la ruta." % ruta)
    with open(ruta, "r", encoding="utf-8", newline="") as fh:
        crudo = fh.read()
    crlf = "\r\n" in crudo
    texto = crudo.replace("\r\n", "\n")
    if MARCA in texto:
        print("Ya estaba aplicado (%s). No se cambió nada." % MARCA)
        return
    faltan = [a for a in ANCLAS if a not in texto]
    if faltan:
        sys.exit("Este app.py no coincide con lo esperado; no toqué nada. Falta/cambió:\n  - " + "\n  - ".join(faltan))
    corte = texto.rfind('\nif __name__ == "__main__":')
    if corte < 0:
        sys.exit('No encontré el bloque if __name__ == "__main__": al final de app.py.')
    nuevo = texto[:corte] + "\n" + BLOQUE.strip("\n") + "\n\n" + texto[corte:]
    try:
        compile(nuevo, str(ruta), "exec")
    except SyntaxError as e:
        sys.exit("El resultado no compila (línea %s: %s). No escribí nada." % (e.lineno, e.msg))
    copia = ruta.with_name(ruta.name + ".bak_descuentos_planes")
    if copia.exists():
        copia = ruta.with_name(ruta.name + ".bak_descuentos_planes_" + time.strftime("%Y%m%d_%H%M%S"))
    shutil.copy2(ruta, copia)
    with open(ruta, "w", encoding="utf-8", newline="") as fh:
        fh.write(nuevo.replace("\n", "\r\n") if crlf else nuevo)
    print("Listo. Descuento y precio final disponibles al crear y editar planes. Archivo: %s" % ruta)
    print("Copia de seguridad: %s" % copia)


if __name__ == "__main__":
    main()
