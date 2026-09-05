import { initializeApp } from "https://www.gstatic.com/firebasejs/10.4.0/firebase-app.js";
import { getDatabase, ref, set, onValue, update, get } from "https://www.gstatic.com/firebasejs/10.4.0/firebase-database.js";

const firebaseConfig = {
    apiKey: "AIzaSyA-x8ZZvJXAOK7Q18PVWPybmfPZ7xDBNHo",
    authDomain: "tablero-pruebas.firebaseapp.com",
    databaseURL: "https://tablero-pruebas-default-rtdb.firebaseio.com",
    projectId: "tablero-pruebas"
};

let app, db;
let listaMotivos = [];
let folioEnEspera = null; 
let folioAEntregar = null;
let califAEntregar = null;

try {
    app = initializeApp(firebaseConfig);
    db = getDatabase(app);
    
    onValue(ref(db, 'citas_diarias'), (snapshot) => {
        renderizarTablas(snapshot.val() || {});
    });

    onValue(ref(db, 'config_prepiking/motivos'), (snapshot) => {
        listaMotivos = snapshot.val() || ["Se le hizo tarde", "Confundió el horario", "Canceló la cita", "Imprevisto personal/ laboral", "Problema de salud"];
        get(ref(db, 'citas_diarias')).then(snap => renderizarTablas(snap.val() || {}));
    });

    onValue(ref(db, 'historial_completado'), (snapshot) => {
        renderizarEntregados(snapshot.val() || {});
    });

} catch (error) {
    console.error("Firebase error: ", error.message);
}

// LOGICA MODAL ACTUALIZACIONES (24 HORAS EXACTAS)
document.addEventListener("DOMContentLoaded", () => {
    let patchTime = localStorage.getItem('patch_v2_time');
    let dismiss = localStorage.getItem('patch_v2_visto');
    const now = Date.now();

    // Si es la primera vez que carga, registramos la hora actual
    if (!patchTime) {
        patchTime = now;
        localStorage.setItem('patch_v2_time', patchTime);
    }

    // Calcula si pasaron menos de 24 horas (86400000 milisegundos)
    if (!dismiss && (now - parseInt(patchTime) < 86400000)) {
        document.getElementById('modalPatchNotes').style.display = 'flex';
    }

    if (document.getElementById('btnCerrarPatch')) {
        document.getElementById('btnCerrarPatch').addEventListener('click', () => {
            localStorage.setItem('patch_v2_visto', 'true');
            document.getElementById('modalPatchNotes').style.display = 'none';
        });
    }
});


// FORMATEADOR DE NOMBRES (Ejem: 01 ASE- LUIS JARED)
function formatearAsesor(nombre) {
    if (!nombre) return 'Sin Asesor';
    if (nombre.includes('[GARANTÍA]')) return nombre; 
    if (nombre.includes('-')) {
        let partes = nombre.split('-');
        let nombreLimpio = partes[1].trim().split(' ');
        return nombreLimpio.slice(0, 2).join(' ');
    }
    return nombre;
}

// PROCESAMIENTO EXCEL (PROTEGIDO)
const inputExcel = document.getElementById('excelFile');
if(inputExcel) {
    inputExcel.addEventListener('click', function() { this.value = null; });
    inputExcel.addEventListener('change', function(e) {
        const file = e.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = function(evt) {
            try {
                const data = new Uint8Array(evt.target.result);
                const workbook = XLSX.read(data, {type: 'array', cellDates: true});
                const sheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonDatos = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ""}); 

                if (!jsonDatos || jsonDatos.length === 0) return;

                let nuevosRegistros = {};
                for (let i = 0; i < jsonDatos.length; i++) {
                    let fila = jsonDatos[i];
                    let estatus = fila[11] ? fila[11].toString().trim().toUpperCase() : "";
                    
                    if (estatus === 'INGRESADA') {
                        let folio = fila[0] ? fila[0].toString().trim() : '';
                        if(folio.endsWith('.0')) folio = folio.slice(0, -2);
                        
                        let fechaRaw = fila[1];
                        let fecha = fechaRaw instanceof Date ? fechaRaw.toLocaleDateString('es-MX') : fila[1];
                        let horaRaw = fila[2];
                        let hora = horaRaw instanceof Date ? horaRaw.toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'}) : fila[2];

                        let placas = fila[6] ? fila[6].toString().trim() : 'S/P';
                        let vin = fila[7] ? fila[7].toString().trim() : 'S/V';
                        let servicio = jsonDatos[i+1]?.[4] || "";
                        let manoObra = jsonDatos[i+1]?.[5] || "";
                        let detalle = (servicio && manoObra) ? `${servicio} | ${manoObra}` : `${servicio}${manoObra}`;
                        
                        nuevosRegistros[folio] = {
                            Folio: folio, Fecha: fecha, Asesor: fila[10] ? fila[10].toString().trim() : 'Sin Asesor',
                            Hora: hora, Cliente: fila[8] ? fila[8].toString().trim() : '',
                            Vehiculo: fila[3] ? fila[3].toString().trim() : '', 
                            VIN: vin, Placas: placas, Servicio: detalle,
                            asistio: "", calif: "", timeAsis: "",
                            motivo: "", reagendo: "", fechaReagendo: "", oculto: false
                        };
                    }
                }
                
                if (db) {
                    get(ref(db, 'citas_diarias')).then(snap => {
                        let dbActual = snap.val() || {};
                        let actualizaciones = {};
                        
                        for (let folio in nuevosRegistros) {
                            if (!dbActual[folio]) { actualizaciones[folio] = nuevosRegistros[folio]; }
                        }
                        
                        if(Object.keys(actualizaciones).length > 0) {
                            update(ref(db, 'citas_diarias'), actualizaciones).then(() => {
                                alert(`✅ Excel cargado. Se agregaron ${Object.keys(actualizaciones).length} citas nuevas.\n(Las citas existentes se protegieron).`);
                            });
                        } else {
                            alert("⚠️ No se agregaron citas. Todos los folios del Excel ya existían y fueron protegidos.");
                        }
                    });
                }
            } catch (err) { console.error("Error Excel: ", err); }
        };
        reader.readAsArrayBuffer(file);
    });
}

// RENDERIZADO DE TABLAS
function renderizarTablas(datos) {
    const tbodyPrepiking = document.getElementById('tbody-prepiking');
    const tbodyNoShow = document.getElementById('tbody-noshow');
    const rol = document.getElementById('userRole') ? document.getElementById('userRole').value : 'asesor';
    
    if(!tbodyPrepiking || !tbodyNoShow) return;
    tbodyPrepiking.innerHTML = ''; tbodyNoShow.innerHTML = '';
    
    const headPrepiking = document.getElementById('thead-prepiking');
    const headNoShow = document.getElementById('thead-noshow');
    
    if (!document.getElementById('th-admin-pre')) {
        if(rol === 'admin') {
            headPrepiking.innerHTML += `<th id="th-admin-pre" style="color:red;">Admin</th>`;
            headNoShow.innerHTML += `<th id="th-admin-no" style="color:red;">Admin</th>`;
        }
    } else if (rol !== 'admin') {
        let thPre = document.getElementById('th-admin-pre');
        let thNo = document.getElementById('th-admin-no');
        if(thPre) thPre.remove();
        if(thNo) thNo.remove();
    }

    if(!datos || Object.keys(datos).length === 0) return;

    let arrCitas = Object.values(datos).sort((a, b) => {
        let dateA = a.Fecha.split('/').reverse().join('-');
        let dateB = b.Fecha.split('/').reverse().join('-');
        let cmpFecha = dateA.localeCompare(dateB);
        if (cmpFecha !== 0) return cmpFecha;
        
        let cmpAsesor = a.Asesor.localeCompare(b.Asesor);
        if (cmpAsesor !== 0) return cmpAsesor;
        return a.Hora.localeCompare(b.Hora);
    });

    arrCitas.forEach(cita => {
        if (cita.oculto) return;

        let f = cita.Folio;
        let rowClass = cita.Asesor.includes('01') ? 'fila-p1' : 'fila-p2';
        if(cita.Asesor.includes('[GARANTÍA]')) rowClass = 'fila-garantia';
        
        let disablePrepiking = (rol === 'citas') ? 'disabled' : ''; 
        let disableNoShow = (rol === 'asesor') ? 'disabled' : ''; 
        let btnAdmin = rol === 'admin' ? `<button class="btn-eliminar-admin" data-folio="${f}" title="Borrar Cita" style="margin-top:5px;">🗑️</button>` : '';

        if (cita.asistio === 'No') {
            let optionsMotivo = `<option value="">Selecciona motivo...</option>` + 
                listaMotivos.map(m => `<option value="${m}" ${cita.motivo === m ? 'selected':''}>${m}</option>`).join('');

            let tr = document.createElement('tr');
            tr.className = rowClass;
            tr.innerHTML = `
                <td><span class="asesor-text">${formatearAsesor(cita.Asesor)}</span><br><button class="btn-deshacer" data-folio="${f}" style="background:transparent; color:#ffcc00; border:1px solid #ffcc00; cursor:pointer;" ${disableNoShow}>↩ Deshacer</button></td>
                <td>${cita.Fecha}</td><td style="color:#888;">${f}</td><td>${cita.Cliente}</td><td><span class="vehiculo-text">${cita.Vehiculo}</span></td>
                <td style="font-size:0.9rem;">${cita.VIN}</td>
                <td><input type="text" class="inp-placas" data-folio="${f}" value="${cita.Placas}" style="width:90px; text-align:center; color:#ffbd2e; border:1px solid #444; background:#111; font-weight:bold;"></td>
                <td>
                    <select ${disableNoShow} class="sel-motivo" data-folio="${f}">${optionsMotivo}</select><br>
                    <input type="text" ${disableNoShow} class="inp-comentario" data-folio="${f}" value="${cita.comentarios || ''}" placeholder="Detalles extra..." style="margin-top:5px;">
                </td>
                <td><select ${disableNoShow} class="sel-reagendo" data-folio="${f}"><option value="">¿Reagendó?</option><option value="Sí" ${cita.reagendo==='Sí'?'selected':''}>Sí</option><option value="No" ${cita.reagendo==='No'?'selected':''}>No</option></select></td>
                <td><input type="date" ${disableNoShow} class="inp-fecha" data-folio="${f}" value="${cita.fechaReagendo || ''}" onclick="this.showPicker()" style="cursor:pointer;"></td>
                ${rol==='admin' ? `<td>${btnAdmin}</td>` : ''}
            `;
            tbodyNoShow.appendChild(tr);
        } else {
            let tr = document.createElement('tr');
            tr.className = rowClass;
            tr.innerHTML = `
                <td><span class="asesor-text">${formatearAsesor(cita.Asesor)}</span></td><td>${cita.Fecha}</td><td><span class="hora-text">${cita.Hora}</span></td>
                <td style="color:#888;">${f}</td><td>${cita.Cliente}</td><td><span class="vehiculo-text">${cita.Vehiculo}</span></td>
                <td style="font-size:0.9rem;">${cita.VIN}</td>
                <td><input type="text" class="inp-placas" data-folio="${f}" value="${cita.Placas}" style="width:90px; text-align:center; color:#ffbd2e; border:1px solid #444; background:#111; font-weight:bold;"></td>
                <td style="font-size:1rem; color:#ccc;">${cita.Servicio}</td>
                <td><select ${disablePrepiking} class="sel-asistencia" data-folio="${f}" data-hora="${cita.Hora}"><option value="">Pendiente...</option><option value="Sí" ${cita.asistio === 'Sí' ? 'selected' : ''}>Sí Asistió</option><option value="No">No Asistió</option></select>
                <span class="timestamp">${cita.timeAsis ? '⏰ ' + cita.timeAsis : ''}</span></td>
                <td><select ${disablePrepiking} class="sel-calif" data-folio="${f}" ${cita.asistio !== 'Sí' ? 'disabled' : ''} style="border-color: ${cita.asistio === 'Sí' ? '#ffcc00' : '#555'};">
                    <option value="">Calificar...</option>
                    <option value="molesto" ${cita.calif==='molesto'?'selected':''}>Cliente Molesto</option>
                    <option value="dudoso" ${cita.calif==='dudoso'?'selected':''}>Cliente Dudoso</option>
                    <option value="contento" ${cita.calif==='contento'?'selected':''}>Cliente Contento</option>
                    <option value="excelente" ${cita.calif==='excelente'?'selected':''}>¡Entrega Excelente!</option>
                </select></td>
                <td><input type="text" class="inp-comentario" data-folio="${f}" value="${cita.comentarios || ''}" placeholder="Escribe aquí..." ${disablePrepiking}></td>
                <td style="text-align:center;">
                    <button class="btn-entregar" data-folio="${f}" style="background:#00ff00; color:#000; border:none; padding:5px 10px; cursor:pointer; font-weight:bold; border-radius:3px; display: ${disablePrepiking ? 'none' : 'inline-block'};">✔️ ENTREGAR</button>
                    <br>${btnAdmin}
                </td>
            `;
            tbodyPrepiking.appendChild(tr);
        }
    });
    asignarEventosDinamicos();
}

function renderizarEntregados(datos) {
    const tbody = document.getElementById('tbody-entregados');
    if(!tbody) return;
    tbody.innerHTML = '';
    
    let rol = document.getElementById('userRole') ? document.getElementById('userRole').value : 'asesor';
    
    const headEntregados = document.getElementById('thead-entregados');
    if (!document.getElementById('th-admin-entre')) {
        if(rol === 'admin') headEntregados.innerHTML += `<th id="th-admin-entre" style="color:red;">Acción</th>`;
    } else if (rol !== 'admin') {
        let thEntre = document.getElementById('th-admin-entre');
        if(thEntre) thEntre.remove();
    }

    let limite7Dias = Date.now() - (7 * 24 * 60 * 60 * 1000);
    let arr = Object.values(datos).filter(cita => cita.timestampEntrega && cita.timestampEntrega >= limite7Dias);
    arr.sort((a,b) => b.timestampEntrega - a.timestampEntrega);
    
    arr.forEach(cita => {
        let tr = document.createElement('tr');
        let colorCalif = cita.calif === 'excelente' ? '#00ff00' : cita.calif === 'molesto' ? '#ff0000' : '#ff9900';
        let btnAdmin = rol === 'admin' ? `<td><button class="btn-eliminar-entrega" data-id="${cita.Folio}_${cita.timestampEntrega}" title="Borrar Entrega" style="background:transparent; border:none; font-size:1.5rem; cursor:pointer;">🗑️</button></td>` : '';
        
        tr.innerHTML = `
            <td>${formatearAsesor(cita.Asesor)}</td>
            <td>${cita.Fecha}</td>
            <td>${cita.Folio}</td>
            <td>${cita.Cliente}</td>
            <td style="color:#00ffff;">${cita.Vehiculo}</td>
            <td style="color:#ffbd2e;">${cita.Placas}</td>
            <td style="color:${colorCalif}; font-weight:bold; text-transform:uppercase;">${cita.calif}</td>
            <td style="color:#aaa;">${new Date(cita.timestampEntrega).toLocaleString()}</td>
            ${btnAdmin}
        `;
        tbody.appendChild(tr);
    });

    // LÓGICA ELIMINAR ENTREGAS (SÓLO ADMIN)
    document.querySelectorAll('.btn-eliminar-entrega').forEach(btn => {
        btn.addEventListener('click', (e) => {
            let id = e.target.dataset.id;
            if(confirm("¿Seguro que deseas borrar permanentemente este registro del historial de entregas? (Esto no afectará la cita original si sigue activa)")) {
                set(ref(db, `historial_completado/${id}`), null);
            }
        });
    });
}

// EVENTOS Y LÓGICA
function asignarEventosDinamicos() {
    document.querySelectorAll('.inp-placas').forEach(el => {
        el.addEventListener('change', (e) => {
            update(ref(db, `citas_diarias/${e.target.dataset.folio}`), { Placas: e.target.value.toUpperCase() });
        });
    });

    document.querySelectorAll('.sel-asistencia').forEach(el => {
        el.addEventListener('change', (e) => {
            let val = e.target.value;
            let folio = e.target.dataset.folio;
            
            if (val === 'No') {
                folioEnEspera = folio;
                document.getElementById('modalNoShow').style.display = 'flex';
                return; 
            }

            let ahora = new Date();
            let horaActualStr = ahora.toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});

            if (val === 'Sí') {
                let horaCita = e.target.dataset.hora;
                if (horaCita) {
                    let [h, m] = horaCita.split(':').map(Number);
                    let fechaCita = new Date();
                    fechaCita.setHours(h, m, 0, 0);
                    
                    if (((ahora - fechaCita) / 60000) >= 20) {
                        let randomGif = Math.floor(Math.random() * 4) + 1;
                        document.getElementById('tardyGif').src = `./${randomGif}.gif`; 
                        document.getElementById('gifOverlay').style.display = 'flex';
                        setTimeout(() => { document.getElementById('gifOverlay').style.display = 'none'; }, 3000);
                    }
                }
            }
            update(ref(db, `citas_diarias/${folio}`), { asistio: val, timeAsis: horaActualStr, calif: "" });
        });
    });

    document.querySelectorAll('.btn-deshacer').forEach(el => {
        el.addEventListener('click', (e) => update(ref(db, `citas_diarias/${e.target.dataset.folio}`), { asistio: "", timeAsis: "" }));
    });

    document.querySelectorAll('.inp-comentario').forEach(el => {
        el.addEventListener('change', (e) => update(ref(db, `citas_diarias/${e.target.dataset.folio}`), { comentarios: e.target.value }));
    });

    document.querySelectorAll('.sel-motivo, .sel-reagendo, .inp-fecha').forEach(el => {
        el.addEventListener('change', (e) => {
            let f = e.target.dataset.folio;
            let act = {};
            act[e.target.classList.contains('sel-motivo') ? 'motivo' : e.target.classList.contains('sel-reagendo') ? 'reagendo' : 'fechaReagendo'] = e.target.value;
            update(ref(db, `citas_diarias/${f}`), act);
        });
    });

    document.querySelectorAll('.btn-eliminar-admin').forEach(el => {
        el.addEventListener('click', (e) => {
            let f = e.target.dataset.folio;
            if(confirm("¿Seguro que deseas borrar permanentemente el Folio " + f + "?")) set(ref(db, `citas_diarias/${f}`), null);
        });
    });

    document.querySelectorAll('.sel-calif').forEach(el => {
        el.addEventListener('change', (e) => {
            update(ref(db, `citas_diarias/${e.target.dataset.folio}`), { calif: e.target.value });
        });
    });

    // ENTREGAR UNIDAD (ABRE MODAL)
    document.querySelectorAll('.btn-entregar').forEach(btn => {
        btn.addEventListener('click', (e) => {
            let folio = e.target.dataset.folio;
            get(ref(db, `citas_diarias/${folio}`)).then(snap => {
                if(snap.exists()) {
                    let citaData = snap.val();
                    if(!citaData.calif || citaData.calif === "") {
                        alert("⚠️ Debes asignar una CALIFICACIÓN antes de entregar la unidad.");
                        return;
                    }
                    folioAEntregar = folio;
                    califAEntregar = citaData.calif;
                    document.getElementById('modalEntregar').style.display = 'flex';
                }
            });
        });
    });
}

// LOGICA MODAL NO SHOW
if(document.getElementById('btnConfirmNoShow')) {
    document.getElementById('btnConfirmNoShow').addEventListener('click', () => {
        if(folioEnEspera) {
            update(ref(db, `citas_diarias/${folioEnEspera}`), { asistio: 'No', timeAsis: new Date().toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'}), calif: "" });
        }
        document.getElementById('modalNoShow').style.display = 'none';
        folioEnEspera = null;
    });
    document.getElementById('btnCancelNoShow').addEventListener('click', () => {
        if(folioEnEspera) {
            let select = document.querySelector(`.sel-asistencia[data-folio="${folioEnEspera}"]`);
            if(select) select.value = "";
        }
        document.getElementById('modalNoShow').style.display = 'none';
        folioEnEspera = null;
    });
}

// LOGICA MODAL ENTREGAR UNIDAD
if(document.getElementById('btnConfirmEntregar')) {
    document.getElementById('btnConfirmEntregar').addEventListener('click', () => {
        if(folioAEntregar && califAEntregar) {
            let val = califAEntregar;
            let folio = folioAEntregar;
            document.getElementById('modalEntregar').style.display = 'none';

            let anim = document.getElementById('ko-anim');
            anim.innerText = val === 'excelente' ? "¡GAME!" : val === 'molesto' ? "¡DANGER!" : val === 'dudoso' ? "¡WARNING!" : "¡K.O.!";
            anim.style.color = val === 'excelente' ? "#00ff00" : val === 'molesto' ? "#ff0000" : val === 'dudoso' ? "#ff9900" : "#ffcc00";
            
            if(val === 'molesto') { document.body.classList.add('screen-shake'); setTimeout(()=>document.body.classList.remove('screen-shake'),500); }
            anim.classList.add('show-ko');

            get(ref(db, `citas_diarias/${folio}`)).then(snap => {
                let citaData = snap.val();
                setTimeout(() => {
                    anim.classList.remove('show-ko');
                    citaData.timestampEntrega = Date.now();
                    set(ref(db, `historial_completado/${folio}_${citaData.timestampEntrega}`), citaData).then(() => {
                        update(ref(db, `citas_diarias/${folio}`), { oculto: true });
                    });
                    folioAEntregar = null;
                    califAEntregar = null;
                }, 1500);
            });
        }
    });
    document.getElementById('btnCancelEntregar').addEventListener('click', () => {
        document.getElementById('modalEntregar').style.display = 'none';
        folioAEntregar = null;
        califAEntregar = null;
    });
}


// ROLES Y FUNCIONES ADMIN
if(document.getElementById('btnUnlock')) {
    document.getElementById('btnUnlock').addEventListener('click', () => {
        let pin = prompt("Ingresa el código de acceso:");
        let select = document.getElementById('userRole');
        
        if (pin === '0520') {
            select.innerHTML = `<option value="asesor">👤 Perfil: Asesor</option><option value="citas">📋 Perfil: Citas</option>`;
            select.value = 'citas'; select.dispatchEvent(new Event('change'));
        } else if (pin === '2099') {
            select.innerHTML = `<option value="asesor">👤 Perfil: Asesor</option><option value="citas">📋 Perfil: Citas</option><option value="admin">⚙️ Perfil: Super Admin</option>`;
            select.value = 'admin'; select.dispatchEvent(new Event('change'));
        } else if (pin) {
            alert("❌ Código incorrecto.");
        }
    });
}

if(document.getElementById('userRole')) {
    document.getElementById('userRole').addEventListener('change', function() {
        let rol = this.value;
        document.getElementById('btnTabNoShow').style.display = (rol === 'citas' || rol === 'admin') ? 'block' : 'none';
        document.getElementById('btnTabEntregados').style.display = (rol === 'admin') ? 'block' : 'none';
        
        document.getElementById('admin-controls').style.display = (rol === 'admin') ? 'flex' : 'none';
        document.getElementById('uploadArea').style.display = (rol === 'citas' || rol === 'admin') ? 'block' : 'none';
        
        document.getElementById('btnGarantia').style.display = (rol === 'citas') ? 'block' : 'none';

        if(rol === 'asesor') document.querySelector('[data-tab="prepiking"]').click();
        get(ref(db, 'citas_diarias')).then(snap => renderizarTablas(snap.val() || {}));
        get(ref(db, 'historial_completado')).then(snap => renderizarEntregados(snap.val() || {}));
    });
}

document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', (e) => {
        document.querySelectorAll('.tab-btn, .table-container').forEach(el => el.classList.remove('active'));
        e.target.classList.add('active');
        document.getElementById(`tab-${e.target.dataset.tab}`).classList.add('active');
    });
});

if(document.getElementById('btnLimpiarBase')) {
    document.getElementById('btnLimpiarBase').addEventListener('click', () => {
        if(confirm("⚠️ ATENCIÓN: ¿Deseas limpiar las citas del día?\n\n(Los NO SHOW y los autos procesados en el TALLER quedarán protegidos).")) {
            let pin = prompt("Ingresa código de Admin (2099):");
            if (pin === '2099') {
                get(ref(db, 'citas_diarias')).then(snap => {
                    let datos = snap.val() || {};
                    let datosRetenidos = {};
                    
                    for(let folio in datos) {
                        if(datos[folio].asistio === 'No' || (datos[folio].estado_taller && datos[folio].estado_taller !== '')) {
                            datosRetenidos[folio] = datos[folio];
                        }
                    }
                    set(ref(db, 'citas_diarias'), datosRetenidos).then(() => {
                        alert("✅ Día reiniciado. Las citas de Taller y No Show están seguras.");
                        document.getElementById('tbody-prepiking').innerHTML = '';
                    });
                });
            } else { alert("❌ Código incorrecto."); }
        }
    });
}

// SIN CITA Y GARANTÍAS INTEGRADO
function inyectarFilaRapida(esGarantia) {
    let rowId = esGarantia ? 'row-garantia' : 'row-sin-cita';
    if(document.getElementById(rowId)) return;
    
    const tbody = document.getElementById('tbody-prepiking');
    const tr = document.createElement('tr');
    tr.id = rowId; 
    tr.style.background = esGarantia ? 'rgba(138, 43, 226, 0.2)' : 'rgba(0, 68, 255, 0.2)';
    
    let hoyHora = new Date().toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
    let hoyFecha = new Date().toISOString().split('T')[0];

    let asesorInput = `<span class="asesor-text" style="color:#fff; font-size:1.5rem;">SIN CITA</span>`;
    if (esGarantia) {
        asesorInput = `
            <select id="${rowId}-asesor" class="inp-comentario" style="font-weight:bold; color:#8a2be2;">
                <option value="[GARANTÍA] LUIS JARED">LUIS JARED</option>
                <option value="[GARANTÍA] EUGENIA NARCIA">EUGENIA NARCIA</option>
                <option value="[GARANTÍA] OTRO ASESOR">OTRO ASESOR</option>
            </select>
        `;
    }

    let horaInput = `<input type="text" id="${rowId}-hora" class="inp-comentario" placeholder="HH:MM" maxlength="5" value="${hoyHora}" oninput="let v=this.value.replace(/[^0-9]/g,''); if(v.length>2) v=v.slice(0,2)+':'+v.slice(2); this.value=v.slice(0,5);">`;

    tr.innerHTML = `
        <td>${asesorInput}</td>
        <td><input type="date" id="${rowId}-fecha" class="inp-comentario" value="${hoyFecha}"></td>
        <td>${horaInput}</td>
        <td><input type="text" id="${rowId}-folio" placeholder="Folio" class="inp-comentario" style="width:70px;"></td>
        <td><input type="text" id="${rowId}-cliente" placeholder="Nombre" class="inp-comentario"></td>
        <td><input type="text" id="${rowId}-vehiculo" placeholder="Auto" class="inp-comentario"></td>
        <td><input type="text" id="${rowId}-vin" placeholder="VIN" class="inp-comentario" style="width:100px;"></td>
        <td><input type="text" id="${rowId}-placas" placeholder="Placas" class="inp-comentario" style="width:80px;"></td>
        <td><input type="text" id="${rowId}-servicio" placeholder="Servicio/Detalle" class="inp-comentario"></td>
        <td colspan="4" style="text-align:center; vertical-align:middle;">
            <button id="btnSave-${rowId}" style="background:${esGarantia ? '#8a2be2' : 'var(--smash-blue)'}; color:#fff; border:1px solid #fff; padding:8px 15px; cursor:pointer; font-family:'Bebas Neue'; font-size:1.2rem; border-radius:3px;">💾 GUARDAR</button>
            <button onclick="document.getElementById('${rowId}').remove()" style="background:#ff0000; color:#fff; border:1px solid #fff; padding:8px 15px; cursor:pointer; font-family:'Bebas Neue'; font-size:1.2rem; margin-left:5px; border-radius:3px;">❌</button>
        </td>
    `;
    tbody.prepend(tr);

    document.getElementById(`btnSave-${rowId}`).addEventListener('click', () => {
        let f = document.getElementById(`${rowId}-folio`).value.trim();
        let h = document.getElementById(`${rowId}-hora`).value.trim();
        let timeRegex = /^([01]\d|2[0-3]):([0-5]\d)$/;

        if(!f) { alert("⚠️ El Folio es obligatorio."); return; }
        if(!timeRegex.test(h)) { alert("⚠️ Formato de hora inválido. Usa HH:MM (24hrs). Ejemplo: 14:30"); return; }
        
        let pDate = document.getElementById(`${rowId}-fecha`).value.split('-');
        let fechaFormato = `${pDate[2]}/${pDate[1]}/${pDate[0]}`;

        let nombreAsesor = esGarantia ? document.getElementById(`${rowId}-asesor`).value : 'SIN CITA';

        let nuevaCita = {
            Folio: f, Fecha: fechaFormato, Asesor: nombreAsesor, Hora: h,
            Cliente: document.getElementById(`${rowId}-cliente`).value.toUpperCase(),
            Vehiculo: document.getElementById(`${rowId}-vehiculo`).value.toUpperCase(),
            VIN: document.getElementById(`${rowId}-vin`).value.toUpperCase() || 'S/V',
            Placas: document.getElementById(`${rowId}-placas`).value.toUpperCase() || 'S/P',
            Servicio: document.getElementById(`${rowId}-servicio`).value.toUpperCase(),
            asistio: "Sí", timeAsis: h, calif: "", comentarios: "",
            motivo: "", reagendo: "", fechaReagendo: "", oculto: false
        };
        update(ref(db, `citas_diarias/${f}`), nuevaCita).then(() => document.getElementById(rowId).remove());
    });
}

if(document.getElementById('btnSinCita')) {
    document.getElementById('btnSinCita').addEventListener('click', () => inyectarFilaRapida(false));
}
if(document.getElementById('btnGarantia')) {
    document.getElementById('btnGarantia').addEventListener('click', () => inyectarFilaRapida(true));
}


// PANEL ADMIN DE MOTIVOS
if(document.getElementById('btnConfigMotivos')) {
    document.getElementById('btnConfigMotivos').addEventListener('click', () => {
        document.getElementById('modalMotivos').style.display = 'flex';
        renderizarListaMotivosAdmin();
    });

    document.getElementById('btnCerrarMotivos').addEventListener('click', () => {
        document.getElementById('modalMotivos').style.display = 'none';
    });

    document.getElementById('btnGuardarNuevoMotivo').addEventListener('click', () => {
        let nuevo = document.getElementById('nuevoMotivoInput').value.trim();
        if(nuevo) {
            listaMotivos.push(nuevo);
            set(ref(db, 'config_prepiking/motivos'), listaMotivos).then(() => {
                document.getElementById('nuevoMotivoInput').value = '';
                renderizarListaMotivosAdmin();
            });
        }
    });
}

function renderizarListaMotivosAdmin() {
    const container = document.getElementById('listaMotivosContainer');
    container.innerHTML = '';
    
    listaMotivos.forEach((motivo, index) => {
        let row = document.createElement('div');
        row.style.display = 'flex'; row.style.justifyContent = 'space-between'; row.style.alignItems = 'center';
        row.style.padding = '10px'; row.style.borderBottom = '1px solid #333';
        
        row.innerHTML = `
            <span style="font-size:1.3rem; color:#fff; font-family:'Teko', sans-serif; letter-spacing:1px;">${motivo}</span>
            <button style="background:#ff0000; border:none; color:#fff; border-radius:3px; cursor:pointer; padding:5px 15px; font-size:1.2rem;" onclick="eliminarMotivo(${index})">🗑️</button>
        `;
        container.appendChild(row);
    });
}

window.eliminarMotivo = function(index) {
    if(confirm("¿Seguro que deseas eliminar este motivo de la lista?")) {
        listaMotivos.splice(index, 1);
        set(ref(db, 'config_prepiking/motivos'), listaMotivos).then(() => {
            renderizarListaMotivosAdmin();
        });
    }
};