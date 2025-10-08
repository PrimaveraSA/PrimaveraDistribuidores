    import { supabase } from "./configDB.js";

    /* === CACHE COINCIDENCIAS BUENAS === */
    let cacheCoincidenciasBuenas = new Map();

    /* === UMBRALES === */
    const UMBRAL_BUENA = 67;
    const UMBRAL_DESCARTE = 0;

    /* === TRACKING DE DUPLICADOS === */
    let vistos = new Set();

    // Variables iniciales para almacenar los nombres originales
    let priceData = [], masterData = null;
    let masterOriginalName = "", priceOriginalName = "";

    let priceCols = { codigo:null, desc:null, precio:null };
    let masterCols = { producto:null, unidad:null, costo:null };
    let reemplazosCosto = new Map();
    let cachePendientes = new Map();

    const log = document.getElementById("log");

    // Normaliza y limpia un string para comparación
    const normalize = str => String(str || "")
    .normalize("NFD")                        // quitar tildes
    .replace(/[\u0300-\u036f]/g, "")         // eliminar marcas de acento
    .replace(/[-*\/().,]/g, "")              // quitar símbolos comunes
    .replace(/\s+/g, " ")                    // unificar espacios
    .replace(/^S\/|^\$|^USD|^US\$|^€|^EUR|^£/i, "") // quitar símbolos de moneda al inicio
    .toUpperCase()
    .trim();

    const ignoreWords = new Set(["DEL","LA","EL","LOS","LAS","Y","EN","CON","PARA","S/"]);


    // Función para generar un timestamp legible
    function timestamp() {
        const now = new Date();
        const pad = n => n.toString().padStart(2, "0");
        return `${now.getFullYear()}${pad(now.getMonth()+1)}${pad(now.getDate())}_${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
    }

    /* === Funciones auxiliares === */
    function cellToString(cell){
        if(cell === null || cell === undefined) return "";
        if(typeof cell === "object"){
            if(cell.richText) return cell.richText.map(r=>r.text).join('');
            if(cell.text) return cell.text;
            if(cell.formula) return cell.result ?? "";
            if(cell.value !== undefined && cell.value !== null) return String(cell.value);
            return String(cell);
        }
        return String(cell);
    }

    function cleanPrice(str){
        str = cellToString(str).replace(/[^\d.,]/g,'').replace(',', '.');
        return parseFloat(str) || 0;
    }

    function renderPreview(containerId, data, maxRows=10){
        const container = document.getElementById(containerId);
        container.innerHTML = '';
        if(!data || data.length<1) return;
        const table = document.createElement("table");
        const thead = document.createElement("thead");
        const headerRow = document.createElement("tr");
        data[0].forEach(h => {
            const th = document.createElement("th");
            th.textContent = h;
            headerRow.appendChild(th);
        });
        thead.appendChild(headerRow);
        table.appendChild(thead);
        const tbody = document.createElement("tbody");
        for(let i=1; i<Math.min(data.length,1+maxRows); i++){
            const tr = document.createElement("tr");
            (data[i] || []).forEach(cell=>{
                const td = document.createElement("td");
                td.textContent = cell;
                tr.appendChild(td);
            });
            tbody.appendChild(tr);
        }
        table.appendChild(tbody);
        container.appendChild(table);
    }

    /* === CACHE COINCIDENCIAS BUENAS === */
    async function precargarCoincidenciasBuenas() {
        const { data, error } = await supabase
            .from("coincidencias_buenas")
            .select("id, producto_excel1, producto_excel2, precio_excel1, precio_excel2, similitud");

        if (error) {
            console.error("Error cargando coincidencias buenas:", error);
            return;
        }

        data.forEach(c => {
            const clave = normalize(c.producto_excel2); // producto del maestro
            cacheCoincidenciasBuenas.set(clave, c);
        });
    }

    // === Función de similitud mejorada con aprendizaje controlado y filtrado seguro ===
    function similarity(a, b) {
        // === 0️⃣ Obtener lista de coincidencias buenas sin romper el Map ===
        const listaBuenas = Array.isArray(cacheCoincidenciasBuenas)
            ? cacheCoincidenciasBuenas
            : Array.from(cacheCoincidenciasBuenas?.values?.() || []);

        // === 1️⃣ Normalizar y limpiar ===
        let textA = normalize(a);
        let textB = normalize(b);

        // === 2️⃣ Generar diccionario de equivalencias aprendido (solo de coincidencias reales y similares) ===
        let diccionarioAprendido = {};

        listaBuenas.forEach(c => {
            if (!c?.producto_excel1 || !c?.producto_excel2) return;

            const p1 = normalize(c.producto_excel1);
            const p2 = normalize(c.producto_excel2);

            // Si son extremadamente distintos, no aprender de ellos
            const similitudBase = simpleSimilarity(p1, p2);
            if (similitudBase < 50) return;

            const palabras1 = p1.split(/\s+/).filter(w => !ignoreWords.has(w));
            const palabras2 = p2.split(/\s+/).filter(w => !ignoreWords.has(w));

            // Aprende solo si longitud y contenido son parecidos
            const len = Math.min(palabras1.length, palabras2.length);
            for (let i = 0; i < len; i++) {
                const w1 = palabras1[i];
                const w2 = palabras2[i];
                if (
                    w1 !== w2 &&
                    w1.length > 2 &&
                    w2.length > 2 &&
                    simpleWordSim(w1, w2) >= 0.8 // solo si son muy parecidas
                ) {
                    diccionarioAprendido[w1] = w2;
                    diccionarioAprendido[w2] = w1;
                }
            }
        });

        // === 3️⃣ Aplicar el aprendizaje a ambos textos ===
        for (const [key, val] of Object.entries(diccionarioAprendido)) {
            const regex = new RegExp(`\\b${key}\\b`, "gi");
            textA = textA.replace(regex, val);
            textB = textB.replace(regex, val);
        }

        // === 4️⃣ Separar palabras y filtrar irrelevantes ===
        let wordsA = textA.split(/\s+/).filter(w => !ignoreWords.has(w));
        let wordsB = textB.split(/\s+/).filter(w => !ignoreWords.has(w));
        if (wordsA.length === 0) wordsA = textA.split(/\s+/);
        if (wordsB.length === 0) wordsB = textB.split(/\s+/);

        // === 5️⃣ Excluir productos ya confirmados (evita reusarlos) ===
        const productosConfirmados = new Set(
            listaBuenas.map(c => normalize(c.producto_excel1))
        );
        if (
            productosConfirmados.has(normalize(a)) ||
            productosConfirmados.has(normalize(b))
        ) {
            return 0;
        }

        // === 6️⃣ Calcular coincidencias ===
        let matches = 0;
        const palabrasContadas = new Set();

        wordsA.forEach(word => {
            if (palabrasContadas.has(word)) return;

            const isNum = /\d/.test(word);
            const matched =
                wordsB.includes(word) ||
                wordsB.some(bw => bw.includes(word) || word.includes(bw)) ||
                (isNum &&
                    wordsB.some(
                        bw => bw.replace(/\D/g, "") === word.replace(/\D/g, "")
                    ));

            if (matched) {
                matches++;
                palabrasContadas.add(word);
            }
        });

        return (matches / wordsA.length) * 100;
    }

    // === Función auxiliar rápida: similitud básica de texto ===
    function simpleSimilarity(t1, t2) {
        const w1 = t1.split(/\s+/);
        const w2 = t2.split(/\s+/);
        let hits = 0;
        w1.forEach(w => {
            if (w2.includes(w)) hits++;
        });
        return (hits / Math.max(w1.length, 1)) * 100;
    }

    // === Función auxiliar rápida: similitud entre palabras ===
    function simpleWordSim(a, b) {
        if (!a || !b) return 0;
        const min = Math.min(a.length, b.length);
        let same = 0;
        for (let i = 0; i < min; i++) if (a[i] === b[i]) same++;
        return same / Math.max(a.length, b.length);
    }




    /* === Acciones BD === */
    async function guardarCoincidenciaBuena(prod1, prod2, p1, p2, sim){
        const { data, error } = await supabase.from("coincidencias_buenas").insert([{ 
            producto_excel1: prod1, 
            producto_excel2: prod2,
            precio_excel1: p1,
            precio_excel2: p2,
            similitud: sim 
        }]).select("id");
        if(error) console.error(error);
        return data ? data[0] : null;
    }

    /* === Guardar coincidencia pendiente evitando duplicados === */
    async function guardarCoincidenciaPendiente(prod1, prod2, p1, p2, sim){
        const clave = normalize(prod1) + "||" + normalize(prod2);

        // ✅ Revisar si ya existe en cache
        if(cachePendientes.has(clave)) return cachePendientes.get(clave);

        // ✅ Revisar si ya existe en BD
        const { data: existe, error: errCheck } = await supabase
            .from("coincidencias_pendientes")
            .select("id, producto_excel1, producto_excel2, precio_excel1, precio_excel2, similitud")
            .eq("producto_excel1", prod1)
            .eq("producto_excel2", prod2)
            .eq("estado", "pendiente")
            .limit(1);

        if(errCheck) console.error(errCheck);
        if(existe && existe.length>0){
            cachePendientes.set(clave, existe[0]);
            return existe[0]; // ya existe, no insertar duplicado
        }

        // ✅ Insertar nuevo
        const { data, error } = await supabase.from("coincidencias_pendientes").insert([{
            producto_excel1: prod1, 
            producto_excel2: prod2,
            precio_excel1: p1,
            precio_excel2: p2,
            similitud: sim,
            estado: "pendiente"
        }]).select("*");

        if(error) console.error(error);
        if(data && data.length>0){
            cachePendientes.set(clave, data[0]);
            return data[0];
        }
        return null;
    }

    /* === Precargar pendientes desde la BD === */
    async function precargarPendientes() {
        const { data, error } = await supabase
            .from("coincidencias_pendientes")
            .select("*")
            .eq("estado","pendiente");

        if(error) return console.error("Error precargando pendientes:", error);

        cachePendientes.clear();
        const tbody = document.querySelector("#tablaPendientes tbody");
        tbody.innerHTML = "";

        data.forEach(row => {
            const key = normalize(row.producto_excel1) + "||" + normalize(row.producto_excel2);
            cachePendientes.set(key, row);

            // Renderizar en tabla
            renderRow(tbody, row, "pendiente", row.id);
        });
    }

    /* === Limpiar pendientes al terminar la comparación === */
    async function limpiarPendientes(){
        const { error } = await supabase
            .from("coincidencias_pendientes")
            .delete()
            .neq("id", 0); // todos los IDs son distintos de 0 → borra todo

        if(error) console.error("Error limpiando pendientes:", error);

        cachePendientes.clear();
        document.querySelector("#tablaPendientes tbody").innerHTML = "";
    }

    async function moverPendienteABuena(id, row){
        await supabase.from("coincidencias_pendientes").delete().eq("id", id);
        await guardarCoincidenciaBuena(row.producto_excel1,row.producto_excel2,row.precio_excel1,row.precio_excel2,row.similitud);
    }

    async function eliminarPendiente(id){
        const { error } = await supabase
            .from("coincidencias_pendientes")
            .delete()
            .eq("id", id);
        if(error) console.error("Error eliminando pendiente:", error);
    }

    /* === Render tablas === */
    function renderRow(tabla, row, tipo, id){
        const tr = document.createElement("tr");
        tr.innerHTML = `
            <td>${row.producto_excel1}</td>
            <td>${row.producto_excel2||"-"}</td>
            <td>${row.precio_excel1||"-"}</td>
            <td>${row.precio_excel2||"-"}</td>
            <td>${row.similitud}%</td>
            <td class="actions"></td>
        `;

        const actions = tr.querySelector(".actions");

        if(tipo==="pendiente"){
            // Acciones para coincidencias pendientes
            let btnVer = document.createElement("button"); 
            btnVer.textContent="✔ Verificar";
            btnVer.onclick = ()=> moverPendienteABuena(id,row).then(()=>tr.remove());

            let btnDel = document.createElement("button"); 
            btnDel.textContent="❌ Eliminar";
            btnDel.onclick = ()=> eliminarPendiente(id).then(()=>tr.remove());

            actions.appendChild(btnVer); 
            actions.appendChild(btnDel);

        } else if(tipo==="buena"){
            // Acción para coincidencias buenas
            let btnAnular = document.createElement("button");
            btnAnular.textContent="❌ Anular";
            btnAnular.onclick = async ()=>{
                // 1. Borrar de la tabla "coincidencias_buenas"
                await eliminarBuena(id);

                // 2. Insertar en "coincidencias_pendientes"
                const nuevo = await guardarCoincidenciaPendiente(
                    row.producto_excel1, 
                    row.producto_excel2, 
                    row.precio_excel1, 
                    row.precio_excel2, 
                    row.similitud
                );

                // 3. Eliminar la fila actual de "Buenas"
                tr.remove();

                // 4. Renderizar en tabla de pendientes (visual)
                renderRow(document.querySelector("#tablaPendientes tbody"), row, "pendiente", nuevo.id);
            };
            actions.appendChild(btnAnular);
        }

        tabla.appendChild(tr);
    }



    /* === Función extra para eliminar coincidencia buena === */
    async function eliminarBuena(id){
        const { error } = await supabase
            .from("coincidencias_buenas")
            .delete()
            .eq("id", id);
        if(error) console.error("Error al eliminar coincidencia buena:", error);
    }

    /* === Render duplicados solo en web === */
    function renderDuplicado(row){
        const tablaDup = document.querySelector("#tablaDuplicados tbody");
        if(!tablaDup) return; // si no existe la tabla, no hace nada

        const tr = document.createElement("tr");
        tr.innerHTML = `
            <td>${row.producto_excel1}</td>
            <td>${row.producto_excel2||"-"}</td>
            <td>${row.precio_excel1||"-"}</td>
            <td>${row.precio_excel2||"-"}</td>
            <td>${row.similitud}%</td>
        `;
        tablaDup.appendChild(tr);
    }

    // Precargar coincidencias buenas al inicio
    precargarCoincidenciasBuenas();


    /* === EVENTOS DE ARCHIVOS === */
    document.getElementById("priceFile").addEventListener("change", async e=>{
        const file = e.target.files[0];
        if(!file) return;
        priceOriginalName = file.name.replace(/\.xlsx$/i,"");
        const data = await file.arrayBuffer();
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(data);
        const ws = wb.worksheets[0];
        priceData = [];
        ws.eachRow((row,rowNumber)=>{ priceData.push(row.values.slice(1).map(cellToString)); });

        priceData = limpiarMonedas(priceData);

        renderPreview("pricePreview", priceData);
        const headers = priceData[0];
        let html = "Selecciona columnas: Código: <select id='priceCodigo'>";
        headers.forEach((h,i)=> html += `<option value='${i+1}'>${h}</option>`);
        html += "</select> Descripción: <select id='priceDesc'>";
        headers.forEach((h,i)=> html += `<option value='${i+1}'>${h}</option>`);
        html += "</select> Precio: <select id='pricePrecio'>";
        headers.forEach((h,i)=> html += `<option value='${i+1}'>${h}</option>`);
        html += "</select>";
        document.getElementById("priceHeaders").innerHTML = html;
        document.getElementById("masterFile").disabled = false;
        log.textContent = `Archivo de precios cargado (${priceData.length-1} registros). Selecciona columnas.`;
        log.style.display = 'inline-block';
    });

    document.getElementById("masterFile").addEventListener("change", async e=>{
        const file = e.target.files[0];
        if(!file) return;
        masterOriginalName = file.name.replace(/\.xlsx$/i,"");
        const data = await file.arrayBuffer();
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(data);
        const ws = wb.worksheets[0];
        masterData = [];
        ws.eachRow((row, rowNumber) => { if(rowNumber >= 2){ masterData.push(row.values.slice(1).map(cellToString)); } });
        masterData = limpiarMonedas(masterData);
        renderPreview("masterPreview", [masterData[0], ...masterData.slice(1)]);
        const masterHeaders = masterData[0];
        let html = "Selecciona columnas: Producto: <select id='masterProd'>";
        masterHeaders.forEach((h,i)=> html += `<option value='${i+1}'>${h}</option>`);
        html += "</select> Unidad: <select id='masterUmed'>";
        masterHeaders.forEach((h,i)=> html += `<option value='${i+1}'>${h}</option>`);
        html += "</select> Costo IGV: <select id='masterCosto'>";
        masterHeaders.forEach((h,i)=> html += `<option value='${i+1}'>${h}</option>`);
        html += "</select>";
        document.getElementById("masterHeaders").innerHTML = html;
        document.getElementById("processBtn").disabled = false;
        log.textContent = `Archivo maestro cargado (${masterData.length-1} registros) . Selecciona columnas.`;
        log.style.display = 'inline-block';
    });


    /* === PROCESAR === */
    document.getElementById("processBtn").addEventListener("click", async ()=>{

        mostrarCargando();

        if(!masterData || priceData.length===0){
            log.innerHTML = "<span style='color:red'>Faltan archivos o selecciones.</span>";
            return;
        }

        await limpiarPendientes(); 
        await precargarPendientes();

        // Columnas seleccionadas
        priceCols.codigo = parseInt(document.getElementById("priceCodigo").value);
        priceCols.desc   = parseInt(document.getElementById("priceDesc").value);
        priceCols.precio = parseInt(document.getElementById("pricePrecio").value);
        masterCols.producto = parseInt(document.getElementById("masterProd").value);
        masterCols.unidad   = parseInt(document.getElementById("masterUmed").value);
        masterCols.costo    = parseInt(document.getElementById("masterCosto").value);

        let cambios=0, noEncontrados=0, duplicados=0;
        let vistosMaestro = new Set(); 
        let usadosPrecio  = new Set(); 
        let paresUsados   = new Set();

        // ⚠️ El maestro empieza en fila 2, así que usamos slice(2)
        for(let rowData of masterData.slice(2)){
            const masterDesc = normalize(rowData[masterCols.producto-1]);
            const unidad = normalize(rowData[masterCols.unidad-1]);
            if(unidad !== "UNIDAD") continue; // solo procesar UNIDAD

            // 🧩 Revisar si ya existe una coincidencia buena guardada
            const productoMaestro = rowData[masterCols.producto-1];
            const existente = cacheCoincidenciasBuenas.get(normalize(productoMaestro));

            if (existente) {
                // Renderizar directamente sin recalcular ni guardar
                renderRow(
                    document.querySelector("#tablaBuenas tbody"),
                    {
                        producto_excel1: existente.producto_excel1,
                        producto_excel2: existente.producto_excel2,
                        precio_excel1: existente.precio_excel1,
                        precio_excel2: existente.precio_excel2,
                        similitud: existente.similitud
                    },
                    "buena",
                    existente.id
                );

                // Mostrar también en la tabla de duplicados (solo vista)
                renderDuplicado({
                    producto_excel1: existente.producto_excel1,
                    producto_excel2: existente.producto_excel2,
                    precio_excel1: existente.precio_excel1,
                    precio_excel2: existente.precio_excel2,
                    similitud: existente.similitud
                });
                duplicados++;

                // También aplicar el reemplazo de costo, como si se hubiera calculado
                reemplazosCosto.set(productoMaestro, existente.precio_excel1);
                continue; // saltar al siguiente producto
            }


            let bestSim = 0, bestMatch = null;

            // Buscar mejor coincidencia en Excel 1 (precios)
            priceData.slice(1).forEach(p=>{
                const precioDesc = normalize(p[priceCols.desc-1]);
                if(usadosPrecio.has(precioDesc)) return; // ya se usó este precio

                const sim = similarity(masterDesc, p[priceCols.desc-1]);
                if(sim > bestSim){ 
                    bestSim = sim; 
                    bestMatch = p; 
                }
            });

            if(bestMatch){
                const precioNuevo = cleanPrice(bestMatch[priceCols.precio-1]); // Excel 1
                const precioViejo = cleanPrice(rowData[masterCols.costo-1]);   // Excel 2
                const keyMaestro  = masterDesc;
                const keyPrecio   = normalize(bestMatch[priceCols.desc-1]);
                const clavePar    = keyMaestro + "||" + keyPrecio;

                if(vistosMaestro.has(keyMaestro) || paresUsados.has(clavePar)){
                    duplicados++;
                    renderDuplicado({
                        producto_excel1: bestMatch[priceCols.desc-1],         // Excel 1 (Precios)
                        producto_excel2: rowData[masterCols.producto-1],      // Excel 2 (Maestro)
                        precio_excel1: precioNuevo,                           // Precio Excel 1
                        precio_excel2: precioViejo,                           // Precio Excel 2
                        similitud: bestSim.toFixed(2)
                    });
                    continue;
                }

                // marcar como usados
                vistosMaestro.add(keyMaestro);
                usadosPrecio.add(keyPrecio);
                paresUsados.add(clavePar);

                if(bestSim >= UMBRAL_BUENA){
                    cambios++;
                    const nueva = await guardarCoincidenciaBuena(
                        bestMatch[priceCols.desc-1],          // Excel 1
                        rowData[masterCols.producto-1],       // Excel 2
                        precioNuevo,                          // Precio Excel 1
                        precioViejo,                          // Precio Excel 2
                        bestSim
                    );

                    reemplazosCosto.set(rowData[masterCols.producto-1], precioNuevo);

                    renderRow(document.querySelector("#tablaBuenas tbody"), {
                        producto_excel1: bestMatch[priceCols.desc-1],
                        producto_excel2: rowData[masterCols.producto-1],
                        precio_excel1: precioNuevo,
                        precio_excel2: precioViejo,
                        similitud: bestSim.toFixed(2)
                    },"buena", nueva.id);
                } else if(bestSim >= UMBRAL_DESCARTE) {
                    noEncontrados++;
                    const nueva = await guardarCoincidenciaPendiente(
                        bestMatch[priceCols.desc-1],          // Excel 1
                        rowData[masterCols.producto-1],       // Excel 2
                        precioNuevo,                          // Precio Excel 1
                        precioViejo,                          // Precio Excel 2
                        bestSim
                    );
                    renderRow(document.querySelector("#tablaPendientes tbody"), {
                        producto_excel1: bestMatch[priceCols.desc-1],
                        producto_excel2: rowData[masterCols.producto-1],
                        precio_excel1: precioNuevo,
                        precio_excel2: precioViejo,
                        similitud: bestSim.toFixed(2)
                    },"pendiente", nueva.id);
                }
            }
        }

        // Refrescar cache después de guardar nuevas coincidencias
        await precargarCoincidenciasBuenas();


        document.getElementById("descargarBtn").disabled = false;

        mostrarFinalizado();

    });


    document.getElementById("descargarBtn").addEventListener("click", async ()=>{
        await descargarExcelMaestro();
    });


    /* === DESCARGAR EXCELS (MAESTRO + PRECIOS) === */
    async function descargarExcelMaestro(){
        const ts = timestamp();

        // Paleta de colores
        const colores = ["C6EFCE","FFF2CC","FFCCE5","CCE5FF","E2EFDA","F4CCCC","D9E1F2","EAD1DC"];
        const colorMap = new Map();
        let colorIndex = 0;
        document.querySelectorAll("#tablaBuenas tbody tr").forEach(tr=>{
            const prodP = normalize(tr.cells[0].textContent);
            const prodM = normalize(tr.cells[1].textContent);
            if(!colorMap.has(prodP) && !colorMap.has(prodM)){
                const color = colores[colorIndex % colores.length];
                colorMap.set(prodP, color);
                colorMap.set(prodM, color);
                colorIndex++;
            }
        });

        // -------------------------
        // 1) Excel Maestro
        // -------------------------
        const wbMaestro = new ExcelJS.Workbook();
        const wsM = wbMaestro.addWorksheet("Maestro");

        wsM.columns = [
            { key: "cod_prod", width: 10.14 },
            { key: "cod_um", width: 10.14 },
            { key: "cod_cost", width: 10.14 },
            { key: "producto", width: 54 },
            { key: "marca", width: 15 },
            { key: "familia", width: 22.29 },
            { key: "linea", width: 15 },
            { key: "u.medid", width: 8 },
            { key: "multip", width: 8 },
            { key: "Costo IGV", width: 12 },
            { key: "Autocalcular", width: 15 },
            { key: "% Minorista", width: 14.14 },
            { key: "% Mayorista", width: 14.14 },
            { key: "% Especial", width: 12.72 }
        ];

        wsM.mergeCells(1,1,1,masterData[0].length);
        wsM.getCell(1,1).value = "LISTAR-PRODUCTO - Sistema Comercial";
        wsM.getCell(1,1).alignment = { horizontal: "center", vertical: "middle" };

        const headerRowM = wsM.addRow(masterData[0]);
        headerRowM.eachCell(cell=> cell.font = { bold: true });

        for(let i=1; i<masterData.length; i++){
            const row = masterData[i].map(val=>{
                if(val===null||val===undefined) return "";
                const num = parseFloat(String(val).replace(",",".")); 
                return !isNaN(num) && isFinite(num) ? num : val;
            });
            const prodName = masterData[i][masterCols.producto-1];
            if(reemplazosCosto.has(prodName)) row[masterCols.costo-1] = reemplazosCosto.get(prodName);

            const excelRow = wsM.addRow(row);
            const desc = normalize(masterData[i][masterCols.producto-1]);
            if(colorMap.has(desc)){
                const color = colorMap.get(desc);
                excelRow.eachCell(cell=> cell.fill={ type:'pattern', pattern:'solid', fgColor:{argb:color} });
            }
        }

        const bufferM = await wbMaestro.xlsx.writeBuffer();
        saveAs(new Blob([bufferM]), `${masterOriginalName}_${ts}.xlsx`);

        // -------------------------
        // 2) Excel Precios
        // -------------------------
        const wbPrecios = new ExcelJS.Workbook();
        const wsP = wbPrecios.addWorksheet("Precios");

        wsP.mergeCells(1,1,1,priceData[0].length);
        wsP.getCell(1,1).value = "LISTAR-PRECIOS - Sistema Comercial";
        wsP.getCell(1,1).alignment = { horizontal:"center", vertical:"middle" };

        const headerRowP = wsP.addRow(priceData[0]);
        headerRowP.eachCell(cell=> cell.font={ bold:true });

        for(let i=1; i<priceData.length; i++){
            const row = priceData[i];
            const excelRow = wsP.addRow(row);
            const desc = normalize(row[priceCols.desc-1]);
            if(colorMap.has(desc)){
                const color = colorMap.get(desc);
                excelRow.eachCell(cell=> cell.fill={ type:'pattern', pattern:'solid', fgColor:{argb:color} });
            }
        }

        const bufferP = await wbPrecios.xlsx.writeBuffer();
        saveAs(new Blob([bufferP]), `${priceOriginalName}_${ts}.xlsx`);
    }

    const modalConexion = document.getElementById("modalConexion");
    const avisoOnline = document.getElementById("conexionRestaurada");

    function mostrarModalConexion(){
    modalConexion.style.display = "flex";
    }
    function cerrarModalConexion(){
    modalConexion.style.display = "none";
    }

    function mostrarAvisoRestaurado(){
    avisoOnline.style.display = "block";
    setTimeout(()=> avisoOnline.style.display = "none", 5000);
    }

    /* Detectar pérdida y recuperación de Internet */
    window.addEventListener("offline", mostrarModalConexion);
    window.addEventListener("online", ()=>{
    cerrarModalConexion();
    mostrarAvisoRestaurado();
    });

    // === BOTÓN NUEVA COMPARACIÓN ===
    const modalReset = document.getElementById("modalReset");
    const mensajeExito = document.getElementById("mensajeExitoReset");
    const resetBtn = document.getElementById("resetBtn");
    const confirmReset = document.getElementById("confirmReset");
    const cancelReset = document.getElementById("cancelReset");

    // Abrir modal
    resetBtn.addEventListener("click", ()=> modalReset.style.display = "flex");

    // Cancelar
    cancelReset.addEventListener("click", ()=> modalReset.style.display = "none");

    // Confirmar limpiar
    confirmReset.addEventListener("click", async () => {
    modalReset.style.display = "none";

    // limpiar pendientes de la BD
    await limpiarPendientes();

    // limpiar tablas
    document.querySelector("#tablaBuenas tbody").innerHTML = "";
    document.querySelector("#tablaPendientes tbody").innerHTML = "";
    document.querySelector("#tablaDuplicados tbody").innerHTML = "";

    // limpiar vistas previas y selects
    document.getElementById("pricePreview").innerHTML = "";
    document.getElementById("masterPreview").innerHTML = "";
    document.getElementById("priceHeaders").innerHTML = "";
    document.getElementById("masterHeaders").innerHTML = "";

    // limpiar mensajes y desactivar botones
    document.getElementById("log").innerHTML = "";
    document.getElementById("descargarBtn").disabled = true;
    document.getElementById("processBtn").disabled = true;

    // resetear los archivos cargados
    document.getElementById("priceFile").value = "";
    document.getElementById("masterFile").value = "";
    document.getElementById("masterFile").disabled = true;

    // limpiar variables globales si existen
    if (typeof masterData !== "undefined") masterData = null;
    if (typeof priceData !== "undefined") priceData = [];
    if (typeof reemplazosCosto !== "undefined") reemplazosCosto.clear?.();
    if (typeof cachePendientes !== "undefined") cachePendientes.clear?.();

    // mostrar mensaje bonito
    mensajeExito.style.display = "block";
    setTimeout(() => mensajeExito.style.display = "none", 4000);
    });

    // === INDICADOR DE CARGA DE COMPARACIONES ===
    const loadingMsg = document.getElementById("loadingComparaciones");

    // Mostrar mensaje "Cargando..."
    function mostrarCargando() {
    loadingMsg.textContent = "⏳ Cargando comparaciones. Sea Paciente...";
    loadingMsg.classList.remove("finalizado");
    loadingMsg.style.display = "block";
    }

    // Mostrar mensaje "Carga Finalizada"
    function mostrarFinalizado() {
    loadingMsg.textContent = "✅ Carga finalizada correctamente";
    loadingMsg.classList.add("finalizado");
    setTimeout(() => {
        loadingMsg.style.display = "none";
    }, 4000);
    }


    function limpiarMonedas(data) {
    if (!data || data.length === 0) return data;

    const monedaRegex = /^\s*(S\/|\$|USD|US\$|€|EUR|£)\s*$/i;
    const simboloInicioRegex = /^\s*(S\/|\$|USD|US\$|€|EUR|£)\s*/i;

    // Paso 1: limpiar los símbolos dentro de las celdas tipo "S/3.3"
    const cleaned = data.map(row =>
        row.map(cell => {
        if (typeof cell === "string") {
            return cell.replace(simboloInicioRegex, "").trim();
        }
        return cell;
        })
    );

    // Paso 2: detectar columnas que contienen SOLO símbolos de moneda
    const numCols = cleaned[0].length;
    const columnasAEliminar = new Set();

    for (let c = 0; c < numCols; c++) {
        let total = 0;
        let simbolos = 0;
        let vacios = 0;

        for (let r = 1; r < cleaned.length; r++) {
        const val = (cleaned[r][c] || "").trim();
        if (val !== "") {
            total++;
            if (monedaRegex.test(val)) simbolos++;
        } else {
            vacios++;
        }
        }

        // Si casi todo son símbolos o celdas vacías → eliminar
        const totalFilas = cleaned.length - 1;
        const ratioSimbolos = simbolos / totalFilas;
        const ratioVacios = vacios / totalFilas;

        if (ratioSimbolos + ratioVacios > 0.9) {
        columnasAEliminar.add(c);
        }
    }

    // Paso 3: reconstruir sin esas columnas
    const result = cleaned.map(row =>
        row.filter((_, idx) => !columnasAEliminar.has(idx))
    );

    return result;
    }