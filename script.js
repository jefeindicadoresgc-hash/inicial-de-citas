import { initializeApp } from "https://www.gstatic.com/firebasejs/10.4.0/firebase-app.js";
import { getDatabase, ref, set, onValue, update, get } from "https://www.gstatic.com/firebasejs/10.4.0/firebase-database.js";

const firebaseConfig = {
    apiKey: "AIzaSyA-x8ZZvJXAOK7Q18PVWPybmfPZ7xDBNHo",
    authDomain: "tablero-pruebas.firebaseapp.com",
    databaseURL: "https://tablero-pruebas-default-rtdb.firebaseio.com",
    projectId: "tablero-pruebas"
};

let app, db;
try {
    app = initializeApp(firebaseConfig);
    db = getDatabase(app);
    onValue(ref(db, 'citas_diarias'), (snapshot) => {
        renderizarTablas(snapshot.val() || {});
    });
} catch (error) {
    console.error("Firebase error: ", error.message);
}

let listaMotivos = ["Se le hizo tarde", "Confundió el horario", "Canceló la cita", "Imprevisto personal/ laboral", "Problema de salud"];
let folioEnEspera = null; 

// PROCESAMIENTO EXCEL
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

                let records = {};
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
                        
                        records[folio] = {
                            Folio: folio, Fecha: fecha, Asesor: fila[10] ? fila[10].toString().trim() : 'Sin Asesor',
                            Hora: hora, Cliente: fila[8] ? fila[8].toString().trim() : '',
                            Vehiculo: fila[3] ? fila[3].toString().trim() : '', 
                            VIN: vin, Placas: placas, Servicio: detalle,
                            asistio: "", calif: "", timeAsis: "", timeCalif: "",
                            motivo: "", reagendo: "", fechaReagendo: "", oculto: false
                        };
                    }
                }
                if (db) set(ref(db, 'citas_diarias'), records);
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
        let cmpAsesor = a.Asesor.localeCompare(b.Asesor);
        if (cmpAsesor !== 0) return cmpAsesor;
        return a.Hora.localeCompare(b.Hora);
    });

    arrCitas.forEach(cita => {
        if (cita.oculto) return;

        let f = cita.Folio;
        let rowClass = cita.Asesor.includes('01') ? 'fila-p1' : 'fila-p2';
        
        let disablePrepiking = (rol === 'citas') ? 'disabled' : ''; 
        let disableNoShow = (rol === 'asesor') ? 'disabled' : ''; 
        let btnAdmin = rol === 'admin' ? `<td><button class="btn-eliminar-admin" data-folio="${f}" title="Borrar Cita">🗑️</button></td>` : '';

        if (cita.asistio === 'No') {
            let optionsMotivo = `<option value="">Selecciona motivo...</option>` + 
                listaMotivos.map(m => `<option value="${m}" ${cita.motivo === m ? 'selected':''}>${m}</option>`).join('');

            let tr = document.createElement('tr');
            tr.className = rowClass;
            tr.innerHTML = `
                <td><span class="asesor-text">${cita.Asesor}</span><br><button class="btn-deshacer" data-folio="${f}" style="background:transparent; color:#ffcc00; border:1px solid #ffcc00; cursor:pointer;" ${disableNoShow}>↩ Deshacer</button></td>
                <td>${cita.Fecha}</td><td style="color:#888;">${f}</td><td>${cita.Cliente}</td><td><span class="vehiculo-text">${cita.Vehiculo}</span></td>
                <td style="font-size:0.9rem;">${cita.VIN}</td><td style="color:#ffbd2e;">${cita.Placas}</td>
                <td>
                    <select ${disableNoShow} class="sel-motivo" data-folio="${f}">${optionsMotivo}</select><br>
                    <input type="text" ${disableNoShow} class="inp-comentario" data-folio="${f}" value="${cita.comentarios || ''}" placeholder="Detalles extra..." style="margin-top:5px;">
                </td>
                <td><select ${disableNoShow} class="sel-reagendo" data-folio="${f}"><option value="">¿Reagendó?</option><option value="Sí" ${cita.reagendo==='Sí'?'selected':''}>Sí</option><option value="No" ${cita.reagendo==='No'?'selected':''}>No</option></select></td>
                <td><input type="date" ${disableNoShow} class="inp-fecha" data-folio="${f}" value="${cita.fechaReagendo || ''}" onclick="this.showPicker()" style="cursor:pointer;"></td>
                ${btnAdmin}
            `;
            tbodyNoShow.appendChild(tr);
        } else {
            let tr = document.createElement('tr');
            tr.className = rowClass;
            tr.innerHTML = `
                <td><span class="asesor-text">${cita.Asesor}</span></td><td>${cita.Fecha}</td><td><span class="hora-text">${cita.Hora}</span></td>
                <td style="color:#888;">${f}</td><td>${cita.Cliente}</td><td><span class="vehiculo-text">${cita.Vehiculo}</span></td>
                <td style="font-size:0.9rem;">${cita.VIN}</td><td style="color:#ffbd2e;">${cita.Placas}</td>
                <td style="font-size:1rem; color:#ccc;">${cita.Servicio}</td>
                <td><select ${disablePrepiking} class="sel-asistencia" data-folio="${f}" data-hora="${cita.Hora}"><option value="">Pendiente...</option><option value="Sí" ${cita.asistio === 'Sí' ? 'selected' : ''}>Sí Asistió</option><option value="No">No Asistió</option></select>
                <span class="timestamp">${cita.timeAsis ? '⏰ ' + cita.timeAsis : ''}</span></td>
                <td><select ${disablePrepiking} class="sel-calif" data-folio="${f}" ${cita.asistio !== 'Sí' ? 'disabled' : ''} style="border-color: ${cita.asistio === 'Sí' ? '#ffcc00' : '#555'};"><option value="">Calificar...</option><option value="molesto">Cliente Molesto</option><option value="dudoso">Cliente Dudoso</option><option value="contento">Cliente Contento</option><option value="excelente">¡Entrega Excelente!</option></select>
                <span class="timestamp"></span></td>
                <td><input type="text" class="inp-comentario" data-folio="${f}" value="${cita.comentarios || ''}" placeholder="Escribe aquí..." ${disablePrepiking}></td>
                ${btnAdmin}
            `;
            tbodyPrepiking.appendChild(tr);
        }
    });
    asignarEventosDinamicos();
}

// EVENTOS Y LÓGICA
function asignarEventosDinamicos() {
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
                        document.getElementById('tardyGif').src = `./gif/${randomGif}.gif`; // Ruta explícita para GitHub
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
            let val = e.target.value; if(!val) return;
            let folio = e.target.dataset.folio;
            let hora = new Date().toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
            
            let anim = document.getElementById('ko-anim');
            anim.innerText = val === 'excelente' ? "¡GAME!" : val === 'molesto' ? "¡DANGER!" : val === 'dudoso' ? "¡WARNING!" : "¡K.O.!";
            anim.style.color = val === 'excelente' ? "#00ff00" : val === 'molesto' ? "#ff0000" : val === 'dudoso' ? "#ff9900" : "#ffcc00";
            if(val === 'molesto') { document.body.classList.add('screen-shake'); setTimeout(()=>document.body.classList.remove('screen-shake'),500); }
            anim.classList.add('show-ko');

            setTimeout(() => {
                anim.classList.remove('show-ko');
                get(ref(db, `citas_diarias/${folio}`)).then(snap => {
                    if(snap.exists()){
                        let citaData = snap.val(); citaData.calif = val; citaData.timeCalif = hora;
                        set(ref(db, `historial_completado/${folio}_${Date.now()}`), citaData);
                        update(ref(db, `citas_diarias/${folio}`), { oculto: true });
                    }
                });
            }, 1500);
        });
    });
}

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
        document.getElementById('admin-controls').style.display = (rol === 'admin') ? 'flex' : 'none';
        document.getElementById('uploadArea').style.display = (rol === 'citas' || rol === 'admin') ? 'block' : 'none';
        if(rol === 'asesor') document.querySelector('[data-tab="prepiking"]').click();
        get(ref(db, 'citas_diarias')).then(snap => renderizarTablas(snap.val() || {}));
    });
}

document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', (e) => {
        document.querySelectorAll('.tab-btn, .table-container').forEach(el => el.classList.remove('active'));
        e.target.classList.add('active');
        document.getElementById(`tab-${e.target.dataset.tab}`).classList.add('active');
    });
});

if(document.getElementById('btnAgregarMotivo')) {
    document.getElementById('btnAgregarMotivo').addEventListener('click', () => {
        let nuevo = prompt("Nuevo motivo de NO SHOW:");
        if(nuevo && nuevo.trim() !== "") { listaMotivos.push(nuevo); get(ref(db, 'citas_diarias')).then(s => renderizarTablas(s.val() || {})); }
    });
}

if(document.getElementById('btnQuitarMotivo')) {
    document.getElementById('btnQuitarMotivo').addEventListener('click', () => {
        let listadoTxt = listaMotivos.map((m, i) => `${i + 1}. ${m}`).join("\n");
        let idx = prompt("Ingresa el NÚMERO del motivo que deseas eliminar:\n\n" + listadoTxt);
        if(idx && !isNaN(idx) && idx > 0 && idx <= listaMotivos.length) {
            listaMotivos.splice(idx - 1, 1); alert("Motivo eliminado."); get(ref(db, 'citas_diarias')).then(s => renderizarTablas(s.val() || {})); 
        }
    });
}

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

// SIN CITA INTEGRADO
if(document.getElementById('btnSinCita')) {
    document.getElementById('btnSinCita').addEventListener('click', () => {
        if(document.getElementById('row-sin-cita')) return;
        const tbody = document.getElementById('tbody-prepiking');
        const tr = document.createElement('tr');
        tr.id = 'row-sin-cita'; tr.style.background = 'rgba(0, 68, 255, 0.2)';
        
        let hora = new Date().toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
        let fecha = new Date().toLocaleDateString('es-MX');

        tr.innerHTML = `
            <td><span class="asesor-text">SIN CITA</span></td>
            <td>${fecha}</td>
            <td><span class="hora-text">${hora}</span></td>
            <td><input type="text" id="inline-folio" placeholder="Folio" class="inp-comentario" style="width:70px;"></td>
            <td><input type="text" id="inline-cliente" placeholder="Nombre" class="inp-comentario"></td>
            <td><input type="text" id="inline-vehiculo" placeholder="Auto" class="inp-comentario"></td>
            <td><input type="text" id="inline-vin" placeholder="VIN" class="inp-comentario" style="width:100px;"></td>
            <td><input type="text" id="inline-placas" placeholder="Placas" class="inp-comentario" style="width:80px;"></td>
            <td><input type="text" id="inline-servicio" placeholder="Servicio a realizar" class="inp-comentario"></td>
            <td colspan="3" style="text-align:center; vertical-align:middle;">
                <button id="btnSaveInline" style="background:var(--smash-blue); color:#fff; border:1px solid #fff; padding:8px 15px; cursor:pointer; font-family:'Bebas Neue'; font-size:1.2rem; border-radius:3px;">💾 GUARDAR</button>
                <button onclick="document.getElementById('row-sin-cita').remove()" style="background:#ff0000; color:#fff; border:1px solid #fff; padding:8px 15px; cursor:pointer; font-family:'Bebas Neue'; font-size:1.2rem; margin-left:5px; border-radius:3px;">❌</button>
            </td>
        `;
        tbody.prepend(tr);

        document.getElementById('btnSaveInline').addEventListener('click', () => {
            let f = document.getElementById('inline-folio').value.trim();
            if(!f) { alert("⚠️ El Folio es obligatorio."); return; }
            let nuevaCita = {
                Folio: f, Fecha: fecha, Asesor: 'SIN CITA', Hora: hora,
                Cliente: document.getElementById('inline-cliente').value.toUpperCase(),
                Vehiculo: document.getElementById('inline-vehiculo').value.toUpperCase(),
                VIN: document.getElementById('inline-vin').value.toUpperCase() || 'S/V',
                Placas: document.getElementById('inline-placas').value.toUpperCase() || 'S/P',
                Servicio: document.getElementById('inline-servicio').value.toUpperCase(),
                asistio: "Sí", timeAsis: hora, calif: "", comentarios: "",
                motivo: "", reagendo: "", fechaReagendo: "", oculto: false
            };
            update(ref(db, `citas_diarias/${f}`), nuevaCita).then(() => document.getElementById('row-sin-cita').remove());
        });
    });
}