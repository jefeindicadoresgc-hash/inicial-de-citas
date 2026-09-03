// [SECCIÓN 1: CONFIGURACIÓN FIREBASE Y PROTECCIÓN DE ERRORES]
import { initializeApp } from "https://www.gstatic.com/firebasejs/10.4.0/firebase-app.js";
import { getDatabase, ref, set, onValue, update, get } from "https://www.gstatic.com/firebasejs/10.4.0/firebase-database.js";

const firebaseConfig = {
    apiKey: "AIzaSyA-x8ZZvJXAOK7Q18PVWPybmfPZ7xDBNHo",
    authDomain: "tablero-pruebas.firebaseapp.com",
    databaseURL: "https://tablero-pruebas-default-rtdb.firebaseio.com",
    projectId: "tablero-pruebas",
    storageBucket: "tablero-pruebas.firebasestorage.app",
    messagingSenderId: "900913447132",
    appId: "1:900913447132:web:fd3b5cc73af4263d69b419"
};

let app, db;
try {
    app = initializeApp(firebaseConfig);
    db = getDatabase(app);
    
    // Escucha en tiempo real (solo si Firebase conectó bien)
    onValue(ref(db, 'citas_diarias'), (snapshot) => {
        renderizarTablas(snapshot.val() || {});
    });
} catch (error) {
    alert("❌ Advertencia: Firebase no conectó. Revisa tu conexión. Error: " + error.message);
}

let listaMotivos = ["Se le hizo tarde", "Confundió el horario", "Canceló la cita", "Imprevisto personal/ laboral", "Problema de salud"];


// [SECCIÓN 2: LECTURA DEL EXCEL A PRUEBA DE FALLOS]
const inputExcel = document.getElementById('excelFile');

// TRUCO VITAL: Limpiar la memoria del botón al hacer clic. 
// Esto te permite subir el mismo Excel 10 veces seguidas sin que el navegador se bloquee.
inputExcel.addEventListener('click', function() {
    this.value = null; 
});

inputExcel.addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) {
        alert("⚠️ El navegador canceló la selección.");
        return;
    }

    alert("✅ 1/4 - Archivo detectado: " + file.name + "\nIniciando lectura...");

    const reader = new FileReader();

    reader.onload = function(evt) {
        alert("✅ 2/4 - Archivo abierto en memoria. Procesando estructura...");
        try {
            // Se usa ArrayBuffer en lugar de BinaryString para mayor compatibilidad
            const data = new Uint8Array(evt.target.result);
            const workbook = XLSX.read(data, {type: 'array', cellDates: true});
            
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            
            // defval: "" fuerza a que las columnas vacías no desaparezcan y desordenen la tabla
            const jsonDatos = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ""}); 

            if (!jsonDatos || jsonDatos.length === 0) {
                alert("❌ Error: El Excel parece estar vacío o corrupto.");
                return;
            }

            alert("✅ 3/4 - Se leyeron " + jsonDatos.length + " filas. Buscando citas 'INGRESADA'...");

            let records = {};
            let contadorCitas = 0;

            for (let i = 0; i < jsonDatos.length; i++) {
                let fila = jsonDatos[i];
                
                // La Columna L en Excel equivale al índice 11 en JavaScript
                let estatus = fila[11] ? fila[11].toString().trim().toUpperCase() : "";
                
                if (estatus === 'INGRESADA') {
                    let folio = fila[0] ? fila[0].toString().trim() : '';
                    if(folio.endsWith('.0')) folio = folio.slice(0, -2);
                    
                    let fechaRaw = fila[1];
                    let fecha = fechaRaw instanceof Date ? fechaRaw.toLocaleDateString('es-MX') : fila[1];
                    
                    let horaRaw = fila[2];
                    let hora = horaRaw instanceof Date ? horaRaw.toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'}) : fila[2];

                    let servicio = jsonDatos[i+1]?.[4] || "";
                    let manoObra = jsonDatos[i+1]?.[5] || "";
                    let detalle = (servicio && manoObra) ? `${servicio} | ${manoObra}` : `${servicio}${manoObra}`;
                    
                    records[folio] = {
                        Folio: folio, Fecha: fecha, Asesor: fila[10] ? fila[10].toString().trim() : 'Sin Asesor',
                        Hora: hora, Cliente: fila[8] ? fila[8].toString().trim() : '',
                        Vehiculo: fila[3] ? fila[3].toString().trim() : '', Servicio: detalle,
                        asistio: "", calif: "", timeAsis: "", timeCalif: "",
                        motivo: "", reagendo: "", fechaReagendo: "", oculto: false
                    };
                    contadorCitas++;
                }
            }

            if (contadorCitas === 0) {
                alert("⚠️ 4/4 FALLO: No se encontró la palabra 'INGRESADA' en la columna L. Revisa tu archivo.");
                return;
            }

            alert("✅ 4/4 - ¡ÉXITO! Se extrajeron " + contadorCitas + " citas. Subiendo a Firebase...");

            if (db) {
                set(ref(db, 'citas_diarias'), records)
                    .then(() => alert('☁️ ¡Tabla actualizada en todos los equipos!'))
                    .catch((err) => alert('❌ Error al subir a Firebase: ' + err.message));
            } else {
                renderizarTablas(records); // Renderizado de emergencia si no hay internet
            }

        } catch (errLectura) {
            alert("❌ Error procesando el Excel: " + errLectura.message);
            console.error(errLectura);
        }
    };

    reader.onerror = function() { alert("❌ El navegador falló al intentar leer tu computadora."); };
    reader.readAsArrayBuffer(file);
});


// [SECCIÓN 3: RENDERIZADO DE TABLAS]
function renderizarTablas(datos) {
    const tbodyPrepiking = document.getElementById('tbody-prepiking');
    const tbodyNoShow = document.getElementById('tbody-noshow');
    const rol = document.getElementById('userRole').value;
    
    tbodyPrepiking.innerHTML = ''; tbodyNoShow.innerHTML = '';
    
    if(!datos || Object.keys(datos).length === 0) return;

    let arrCitas = Object.values(datos).sort((a, b) => a.Asesor.localeCompare(b.Asesor));

    arrCitas.forEach(cita => {
        if (cita.oculto) return;

        let f = cita.Folio;
        let rowClass = cita.Asesor.includes('01') ? 'fila-p1' : 'fila-p2';
        let disableAsesor = (rol === 'admin' || rol === 'asesor') ? '' : 'disabled';
        let disableCitas = (rol === 'asesor') ? 'disabled' : '';

        if (cita.asistio === 'No') {
            let optionsMotivo = `<option value="">Selecciona motivo...</option>` + 
                listaMotivos.map(m => `<option value="${m}" ${cita.motivo === m ? 'selected':''}>${m}</option>`).join('');

            let tr = document.createElement('tr');
            tr.className = rowClass;
            tr.innerHTML = `
                <td><span class="asesor-text">${cita.Asesor}</span><br><button class="btn-deshacer" data-folio="${f}" style="background:transparent; color:#ffcc00; border:1px solid #ffcc00; cursor:pointer;">↩ Deshacer</button></td>
                <td>${cita.Fecha}</td><td style="color:#888;">${f}</td><td>${cita.Cliente}</td><td><span class="vehiculo-text">${cita.Vehiculo}</span></td>
                <td><select ${disableCitas} class="sel-motivo" data-folio="${f}">${optionsMotivo}</select></td>
                <td><select ${disableCitas} class="sel-reagendo" data-folio="${f}"><option value="">¿Reagendó?</option><option value="Sí" ${cita.reagendo==='Sí'?'selected':''}>Sí</option><option value="No" ${cita.reagendo==='No'?'selected':''}>No</option></select></td>
                <td><input type="date" ${disableCitas} class="inp-fecha" data-folio="${f}" value="${cita.fechaReagendo || ''}"></td>
            `;
            tbodyNoShow.appendChild(tr);
        } else {
            let tr = document.createElement('tr');
            tr.className = rowClass;
            tr.innerHTML = `
                <td><span class="asesor-text">${cita.Asesor}</span></td><td>${cita.Fecha}</td><td><span class="hora-text">${cita.Hora}</span></td>
                <td style="color:#888;">${f}</td><td>${cita.Cliente}</td><td><span class="vehiculo-text">${cita.Vehiculo}</span></td>
                <td style="font-size:1rem; color:#ccc;">${cita.Servicio}</td>
                <td><select ${disableAsesor} class="sel-asistencia" data-folio="${f}"><option value="">Pendiente...</option><option value="Sí" ${cita.asistio === 'Sí' ? 'selected' : ''}>Sí Asistió</option><option value="No">No Asistió</option></select>
                <span class="timestamp">${cita.timeAsis ? '⏰ ' + cita.timeAsis : ''}</span></td>
                <td><select ${disableAsesor} class="sel-calif" data-folio="${f}" ${cita.asistio !== 'Sí' ? 'disabled' : ''} style="border-color: ${cita.asistio === 'Sí' ? '#ffcc00' : '#555'};"><option value="">Calificar...</option><option value="molesto">Cliente Molesto</option><option value="dudoso">Cliente Dudoso</option><option value="contento">Cliente Contento</option><option value="excelente">¡Entrega Excelente!</option></select>
                <span class="timestamp"></span></td>
            `;
            tbodyPrepiking.appendChild(tr);
        }
    });
    asignarEventosDinamicos();
}


// [SECCIÓN 4: LÓGICA DE EVENTOS DINÁMICOS]
function asignarEventosDinamicos() {
    document.querySelectorAll('.sel-asistencia').forEach(el => {
        el.addEventListener('change', (e) => {
            let val = e.target.value;
            if (val === 'No' && !confirm("⚠️ ¿Seguro que el cliente NO ASISTIÓ?")) { e.target.value = ""; return; }
            update(ref(db, `citas_diarias/${e.target.dataset.folio}`), {
                asistio: val, timeAsis: new Date().toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'}), calif: "" 
            });
        });
    });

    document.querySelectorAll('.btn-deshacer').forEach(el => {
        el.addEventListener('click', (e) => update(ref(db, `citas_diarias/${e.target.dataset.folio}`), { asistio: "", timeAsis: "" }));
    });

    document.querySelectorAll('.sel-motivo, .sel-reagendo, .inp-fecha').forEach(el => {
        el.addEventListener('change', (e) => {
            let f = e.target.dataset.folio;
            let act = {};
            act[e.target.classList.contains('sel-motivo') ? 'motivo' : e.target.classList.contains('sel-reagendo') ? 'reagendo' : 'fechaReagendo'] = e.target.value;
            update(ref(db, `citas_diarias/${f}`), act);
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


// [SECCIÓN 5: INTERFAZ Y PESTAÑAS]
document.getElementById('userRole').addEventListener('change', function() {
    let rol = this.value;
    document.getElementById('btnTabNoShow').style.display = (rol === 'citas' || rol === 'admin') ? 'block' : 'none';
    document.getElementById('admin-controls').style.display = (rol === 'admin') ? 'block' : 'none';
    if(rol === 'asesor') document.querySelector('[data-tab="prepiking"]').click();
    get(ref(db, 'citas_diarias')).then(snap => renderizarTablas(snap.val() || {}));
});

document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', (e) => {
        document.querySelectorAll('.tab-btn, .table-container').forEach(el => el.classList.remove('active'));
        e.target.classList.add('active');
        document.getElementById(`tab-${e.target.dataset.tab}`).classList.add('active');
    });
});