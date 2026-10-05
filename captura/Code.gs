// Google Apps Script: recibe las postulaciones del formulario y las guarda en una hoja de cálculo.
// 1) Crea una hoja de cálculo de Google. 2) Extensiones > Apps Script, pega este código.
// 3) Implementar > Nueva implementación > Aplicación web; ejecutar como: tú; acceso: cualquier usuario.
// 4) Copia la URL que termina en /exec y pégala en la constante ENDPOINT de index.html.
const CAMPOS = ['nro_postulante','nombre','email','genero','rango_edad','residencia','experiencia','nivel_profesional','id_cargo','cargo','canal','fecha_aplicacion','creado','estado','etapa_alcanzada'];
const LISTAS = {
  genero: ['Femenino','Masculino','Indefinido'],
  rango_edad: ['18-24','25-34','35-44','45-54','55-64','65+'],
  residencia: ['La Paz','Cochabamba','Santa Cruz de la Sierra','Oruro','Tarija','Sucre','Potosí','Trinidad','Cobija'],
  experiencia: ['< 1 año','2-4 años','>4 años'],
  nivel_profesional: ['Junior','Medio','Senior'],
  id_cargo: ['C1','C2','C3','C4','C6','C7','C8','C9'],
  canal: ['facebook','LinkedIn','Bolsa de trabajo']
};
function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    for (const k in LISTAS) if (LISTAS[k].indexOf(d[k]) < 0) return out({ok:false, error:'Valor fuera de lista: ' + k});
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(d.email || '')) return out({ok:false, error:'Correo no válido'});
    if (!d.nombre || !d.consentimiento) return out({ok:false, error:'Faltan campos obligatorios'});
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    if (sh.getLastRow() === 0) sh.appendRow(CAMPOS);
    d.nro_postulante = 'P' + (40 + sh.getLastRow());   // P41 en adelante: los 40 primeros son del archivo auditado
    sh.appendRow(CAMPOS.map(function(c){ return d[c] === undefined ? '' : d[c]; }));
    return out({ok:true, nro_postulante:d.nro_postulante});
  } catch (err) { return out({ok:false, error:String(err)}); }
  finally { lock.releaseLock(); }
}
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
