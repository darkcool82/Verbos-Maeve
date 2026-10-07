/* =====================================
   PARTE 1 - VARIABLES GLOBALES
===================================== */

let workbook = null;

let verbos = [];
let testActual = [];
let errores = [];
let resultados = [];

let historico =
    JSON.parse(
        localStorage.getItem("historico")
    ) || [];

let temporizador = null;
let segundosRestantes = 0;

let estadisticas =
    JSON.parse(
        localStorage.getItem("estadisticasVerbos")
    ) || {};

let deferredPrompt = null;


/* =====================================
   INICIO
===================================== */

document.addEventListener(
    "DOMContentLoaded",
    async () => {

        registrarEventos();

        cargarHistorico();

        await cargarExcelPorDefecto();

    }
);


/* =====================================
   PARTE 2 - REGISTRO DE EVENTOS
===================================== */

function registrarEventos() {

    const eventos = {

        btnImportar: abrirSelectorExcel,

        btnComenzar: iniciarTest,

        btnEstudiar: estudiarTodos,

        btnRepetirErrores: repetirErrores,

        btnExportarErrores: exportarErrores,

        btnExportarResultados: exportarResultados,

        btnHistorico: mostrarHistorico,

        btnEstadisticas: mostrarEstadisticas,

        btnLimpiarHistorico: limpiarHistorico

    };


    Object.entries(eventos).forEach(
        ([id, funcion]) => {

            const elemento =
                document.getElementById(id);

            if (elemento) {

                elemento.addEventListener(
                    "click",
                    funcion
                );

            }

        }
    );


    const btnLimpiarEstadisticas =
        document.getElementById(
            "btnLimpiarEstadisticas"
        );

    if (btnLimpiarEstadisticas) {

        btnLimpiarEstadisticas.addEventListener(
            "click",
            limpiarEstadisticas
        );

    }

    const inputExcel = document.getElementById("excelFile");

    if (inputExcel) {
        inputExcel.addEventListener("change", importarExcel);
    }
}


/* =====================================
   ABRIR SELECTOR DE EXCEL
===================================== */

function abrirSelectorExcel() {

    const input = document.getElementById("excelFile");

    if (!input) {
        alert("No se encontró el selector de archivos.");
        return;
    }

    input.value = "";
    input.click();
}


/* =====================================
   PARTE 3 - CARGA AUTOMÁTICA DEL EXCEL
===================================== */

async function cargarExcelPorDefecto() {

    try {

        const respuesta =
            await fetch(
                "data/Verbos_ingles.xlsx"
            );

        if (!respuesta.ok) {

            throw new Error(
                "No se encontró el Excel."
            );

        }

        const buffer =
            await respuesta.arrayBuffer();

        workbook =
            XLSX.read(buffer, {
                type: "array"
            });

        generarSelectorHojas();

        console.log(
            "Excel cargado correctamente."
        );

    }
    catch (error) {

        console.error(
            "Error cargando el Excel:",
            error
        );

        console.log(
            "No se ha encontrado el Excel por defecto."
        );

    }
}


/* =====================================
   PARTE 4 - IMPORTAR OTRO EXCEL
===================================== */

async function importarExcel() {

    const input =
        document.getElementById(
            "excelFile"
        );

    if (!input) {

        alert(
            "No se encontró el selector de archivos."
        );

        return;
    }


    const archivo =
        input.files[0];


    if (!archivo) {

        alert(
            "Seleccione un archivo Excel."
        );

        return;
    }


    try {

        const buffer =
            await archivo.arrayBuffer();

        workbook =
            XLSX.read(buffer, {
                type: "array"
            });

        generarSelectorHojas();

        alert(
            "Excel cargado correctamente."
        );

    }
    catch (error) {

        console.error(
            "Error importando Excel:",
            error
        );

        alert(
            "No se pudo cargar el archivo Excel."
        );

    }
}


/* =====================================
   PARTE 5 - SELECTOR DINÁMICO DE HOJAS
===================================== */

function generarSelectorHojas() {

    const contenedor =
        document.getElementById(
            "listaHojas"
        );


    if (!contenedor || !workbook) {

        return;

    }


    contenedor.innerHTML = "";


    workbook.SheetNames.forEach(
        nombre => {

            const item =
                document.createElement(
                    "div"
                );

            item.className =
                "hoja-item";


            const label =
                document.createElement(
                    "label"
                );


            const checkbox =
                document.createElement(
                    "input"
                );

            checkbox.type =
                "checkbox";

            checkbox.className =
                "hojaCheck";

            checkbox.value =
                nombre;

            checkbox.addEventListener("change", actualizarNumeroPreguntasPorSeleccion);

            label.appendChild(
                checkbox
            );

            label.appendChild(
                document.createTextNode(
                    ` ${nombre}`
                )
            );


            item.appendChild(
                label
            );


            contenedor.appendChild(
                item
            );

        }
    );

    actualizarNumeroPreguntasPorSeleccion();
}


function actualizarNumeroPreguntasPorSeleccion() {

    const inputCantidad = document.getElementById("cantidadPreguntas");
    const checks = document.querySelectorAll(".hojaCheck:checked");

    if (!inputCantidad || !workbook) {
        return;
    }

    let total = 0;

    Array.from(checks).forEach(checkbox => {
        const hoja = workbook.Sheets[checkbox.value];
        if (!hoja) return;

        const filas = XLSX.utils.sheet_to_json(hoja, { defval: "" });
        total += filas.filter(fila => String(fila["INFINITIVE"] || "").trim() !== "").length;
    });

    inputCantidad.value = total > 0 ? total : "";
    inputCantidad.max = total > 0 ? total : "";

    const resumen = document.getElementById("resumenSeleccion");
    if (resumen) {
        resumen.textContent = total > 0
            ? `Verbos seleccionados: ${total}`
            : "No hay hojas seleccionadas.";
    }
}


/* =====================================
   PARTE 6 - MODO ESTUDIO
===================================== */

function estudiarTodos() {

    if (!workbook) {

        alert(
            "No hay ningún Excel cargado."
        );

        return;
    }


    const hojas =
        obtenerHojasSeleccionadas();


    if (hojas.length === 0) {

        alert(
            "Seleccione al menos una hoja."
        );

        return;
    }


    const listaVerbos =
        cargarVerbosSeleccionados();


    if (listaVerbos.length === 0) {

        alert(
            "No se encontraron verbos en las hojas seleccionadas."
        );

        return;
    }


    mostrarModoEstudio(
        listaVerbos,
        hojas
    );
}


/* =====================================
   PARTE 7 - OBTENER HOJAS SELECCIONADAS
===================================== */

function obtenerHojasSeleccionadas() {

    const checks =
        document.querySelectorAll(
            ".hojaCheck:checked"
        );


    return Array.from(checks)
        .map(
            checkbox =>
                checkbox.value
        );
}


/* =====================================
   PARTE 8 - LEER VERBOS DE UNA HOJA
===================================== */

function obtenerVerbosHoja(
    nombreHoja
) {

    if (!workbook) {

        return [];

    }


    const hoja =
        workbook.Sheets[
            nombreHoja
        ];


    if (!hoja) {

        return [];

    }


    const filas =
        XLSX.utils.sheet_to_json(
            hoja,
            {
                defval: ""
            }
        );


    return filas
        .map(fila => ({

            infinitivo:
                String(
                    fila["INFINITIVE"] || ""
                ).trim(),

            pasado:
                String(
                    fila["PAST SIMPLE"] || ""
                ).trim(),

            participio:
                String(
                    fila["PAST PARTICIPLE"] || ""
                ).trim(),

            meaning:
                String(
                    fila["MEANING"] || ""
                ).trim()

        }))
        .filter(
            verbo =>
                verbo.infinitivo !== ""
        );
}


/* =====================================
   PARTE 9 - COMBINAR VARIAS HOJAS
===================================== */

function cargarVerbosSeleccionados() {

    const hojas =
        obtenerHojasSeleccionadas();


    let lista = [];


    hojas.forEach(
        hoja => {

            lista =
                lista.concat(
                    obtenerVerbosHoja(
                        hoja
                    )
                );

        }
    );


    return lista;
}


/* =====================================
   PARTE 10 - MODO ESTUDIO
===================================== */

function mostrarModoEstudio(
    listaVerbos,
    hojas
) {

    detenerTemporizador();


    const contenedor =
        document.getElementById(
            "contenido"
        );


    if (!contenedor) {

        return;

    }


    let html = `

        <div class="panel">

            <h2>
                📘 Modo estudio
            </h2>

            <p>
                Hojas:
                ${hojas.join(", ")}
            </p>

            <p>
                Verbos:
                ${listaVerbos.length}
            </p>

        </div>

    `;


    html += `
        <div class="tabla-estudio">
    `;


    listaVerbos.forEach(
        verbo => {

            html += `

                <div class="fila-estudio">

                    <div>
                        ${escapeHTML(verbo.infinitivo)}
                    </div>

                    <div>
                        ${escapeHTML(verbo.pasado)}
                    </div>

                    <div>
                        ${escapeHTML(verbo.participio)}
                    </div>

                    <div>
                        ${escapeHTML(verbo.meaning)}
                    </div>

                </div>

            `;

        }
    );


    html += `
        </div>
    `;


    contenedor.innerHTML =
        html;


    limpiarResultado();
}


/* =====================================
   PARTE 11 - INICIO DEL TEST
===================================== */

function iniciarTest() {

    detenerTemporizador();


    verbos =
        cargarVerbosSeleccionados();


    if (verbos.length === 0) {

        alert(
            "Seleccione al menos una hoja."
        );

        return;
    }


    const inputCantidad =
        document.getElementById(
            "cantidadPreguntas"
        );


    let cantidad =
        inputCantidad
            ? parseInt(
                inputCantidad.value,
                10
            )
            : verbos.length;


    if (
        isNaN(cantidad) ||
        cantidad <= 0
    ) {

        cantidad =
            verbos.length;

    }


    cantidad =
        Math.min(
            cantidad,
            verbos.length
        );


    testActual =
        mezclarArray(
            verbos
        ).slice(
            0,
            cantidad
        );


    errores = [];

    resultados = [];


    renderizarPreguntas();

    iniciarTemporizadorSiProcede();
}


/* =====================================
   PARTE 12 - REPETIR ERRORES
===================================== */

function repetirErrores() {

    detenerTemporizador();


    if (errores.length === 0) {

        alert(
            "No hay errores para repetir."
        );

        return;
    }


    testActual =
        mezclarArray(
            errores
        );


    errores = [];

    resultados = [];


    renderizarPreguntas();

    iniciarTemporizadorSiProcede();


    window.scrollTo({

        top: 0,

        behavior: "smooth"

    });
}


/* =====================================
   PARTE 13 - RENDERIZAR PREGUNTAS
===================================== */

function renderizarPreguntas() {

    const contenedor =
        document.getElementById(
            "contenido"
        );


    if (!contenedor) {

        return;

    }


    let html = "";


    testActual.forEach(
        (verbo, i) => {

            html += `

                <div class="pregunta">

                    <h3>
                        ${escapeHTML(verbo.meaning)}
                    </h3>

                    <div class="fila-verbos">

                        <div class="campo">

                            <input
                                type="text"
                                id="inf_${i}"
                                placeholder="INFINITIVE"
                                autocomplete="off"
                            >

                            <div
                                id="error_inf_${i}"
                                class="error-correccion">
                            </div>

                        </div>


                        <div class="campo">

                            <input
                                type="text"
                                id="pas_${i}"
                                placeholder="PAST SIMPLE"
                                autocomplete="off"
                            >

                            <div
                                id="error_pas_${i}"
                                class="error-correccion">
                            </div>

                        </div>


                        <div class="campo">

                            <input
                                type="text"
                                id="par_${i}"
                                placeholder="PAST PARTICIPLE"
                                autocomplete="off"
                            >

                            <div
                                id="error_par_${i}"
                                class="error-correccion">
                            </div>

                        </div>

                    </div>

                </div>

            `;

        }
    );


    html += `

        <button
            id="btnCorregir"
            type="button">

            Corregir

        </button>

    `;


    contenedor.innerHTML =
        html;


    const btnCorregir =
        document.getElementById(
            "btnCorregir"
        );


    if (btnCorregir) {

        btnCorregir.addEventListener(
            "click",
            corregir
        );

    }


    limpiarResultado();
}


/* =====================================
   PARTE 14 - TEMPORIZADOR
===================================== */

function iniciarTemporizadorSiProcede() {

    detenerTemporizador();


    const checkbox =
        document.getElementById(
            "usarTemporizador"
        );


    if (
        !checkbox ||
        !checkbox.checked
    ) {

        actualizarTemporizador(
            false
        );

        return;
    }


    segundosRestantes =
        testActual.length * 15;


    actualizarTemporizador(
        false
    );


    temporizador =
        setInterval(
            () => {

                segundosRestantes--;

                actualizarTemporizador(
                    false
                );


                if (
                    segundosRestantes <= 0
                ) {

                    detenerTemporizador();

                    corregir();

                }

            },
            1000
        );
}


function actualizarTemporizador(
    mostrarCero = true
) {

    const div =
        document.getElementById(
            "temporizador"
        );


    if (!div) {

        return;

    }


    if (
        !mostrarCero &&
        segundosRestantes <= 0
    ) {

        div.innerHTML =
            "⏱️ 0s";

        return;

    }


    div.innerHTML =
        `⏱️ ${Math.max(
            segundosRestantes,
            0
        )}s`;
}


function detenerTemporizador() {

    if (temporizador !== null) {

        clearInterval(
            temporizador
        );

        temporizador = null;

    }


    segundosRestantes = 0;


    const div =
        document.getElementById(
            "temporizador"
        );


    if (div) {

        div.innerHTML = "";

    }
}


/* =====================================
   PARTE 15 - CORREGIR TEST
===================================== */

function corregir() {

    if (
        testActual.length === 0
    ) {

        return;

    }


    detenerTemporizador();


    errores = [];

    resultados = [];


    let aciertos = 0;


    testActual.forEach(
        (verbo, i) => {

            const infInput =
                document.getElementById(
                    `inf_${i}`
                );

            const pasInput =
                document.getElementById(
                    `pas_${i}`
                );

            const parInput =
                document.getElementById(
                    `par_${i}`
                );


            if (
                !infInput ||
                !pasInput ||
                !parInput
            ) {

                return;

            }


            const infError =
                document.getElementById(
                    `error_inf_${i}`
                );

            const pasError =
                document.getElementById(
                    `error_pas_${i}`
                );

            const parError =
                document.getElementById(
                    `error_par_${i}`
                );


            if (infError) {
                infError.innerHTML = "";
            }

            if (pasError) {
                pasError.innerHTML = "";
            }

            if (parError) {
                parError.innerHTML = "";
            }


            infInput.className = "";
            pasInput.className = "";
            parInput.className = "";


            let verboCorrecto = true;


            /* =====================
               INFINITIVE
            ===================== */

            const infUsuario =
                normalizarTexto(
                    infInput.value
                );


            const infCorrecto =
                normalizarTexto(
                    verbo.infinitivo
                );


            if (
                infUsuario ===
                infCorrecto
            ) {

                infInput.classList.add(
                    "correcto-input"
                );

            }
            else {

                verboCorrecto = false;

                infInput.classList.add(
                    "incorrecto-input"
                );


                if (infError) {

                    infError.innerHTML =
                        escapeHTML(
                            verbo.infinitivo
                        );

                }

            }


            /* =====================
               PAST SIMPLE
            ===================== */

            const pasUsuario =
                normalizarTexto(
                    pasInput.value
                );


            const pasCorrecto =
                normalizarTexto(
                    verbo.pasado
                );


            if (
                pasUsuario ===
                pasCorrecto
            ) {

                pasInput.classList.add(
                    "correcto-input"
                );

            }
            else {

                verboCorrecto = false;

                pasInput.classList.add(
                    "incorrecto-input"
                );


                if (pasError) {

                    pasError.innerHTML =
                        escapeHTML(
                            verbo.pasado
                        );

                }

            }


            /* =====================
               PAST PARTICIPLE
            ===================== */

            const parUsuario =
                normalizarTexto(
                    parInput.value
                );


            const parCorrecto =
                normalizarTexto(
                    verbo.participio
                );


            if (
                parUsuario ===
                parCorrecto
            ) {

                parInput.classList.add(
                    "correcto-input"
                );

            }
            else {

                verboCorrecto = false;

                parInput.classList.add(
                    "incorrecto-input"
                );


                if (parError) {

                    parError.innerHTML =
                        escapeHTML(
                            verbo.participio
                        );

                }

            }


            /* =====================
               ESTADÍSTICAS
            ===================== */

            registrarPregunta(
                verbo
            );


            if (verboCorrecto) {

                aciertos++;

            }
            else {

                errores.push(
                    verbo
                );


                registrarErrorVerbo(
                    verbo
                );

            }


            resultados.push({

                INFINITIVE:
                    verbo.infinitivo,

                "PAST SIMPLE":
                    verbo.pasado,

                "PAST PARTICIPLE":
                    verbo.participio,

                MEANING:
                    verbo.meaning,

                ACIERTO:
                    verboCorrecto
                        ? "SI"
                        : "NO"

            });

        }
    );


    mostrarResultadoFinal(
        aciertos
    );
}


/* =====================================
   PARTE 16 - RESULTADO FINAL
===================================== */

function mostrarResultadoFinal(
    aciertos
) {

    const total =
        testActual.length;


    if (total === 0) {

        return;

    }


    const erroresCantidad =
        total - aciertos;


    const porcentaje =
        Math.round(
            (
                aciertos * 100
            ) /
            total
        );


    const nota =
        (
            aciertos * 10
        ) /
        total;


    const notaFinal =
        nota.toFixed(1);


    const resultado =
        document.getElementById(
            "resultado"
        );


    if (resultado) {

        resultado.innerHTML = `

            <div class="resultado-box">

                <h2>
                    Nota: ${notaFinal}/10
                </h2>

                <p>
                    ✅ Aciertos:
                    ${aciertos}
                </p>

                <p>
                    ❌ Errores:
                    ${erroresCantidad}
                </p>

                <p>
                    📊 Porcentaje:
                    ${porcentaje}%
                </p>

            </div>

        `;

    }


    guardarResultadoHistorico(
        notaFinal,
        porcentaje,
        aciertos,
        erroresCantidad,
        total
    );


    window.scrollTo({

        top: 0,

        behavior: "smooth"

    });
}


/* =====================================
   PARTE 17 - HISTÓRICO
===================================== */

function guardarResultadoHistorico(
    nota,
    porcentaje,
    aciertos,
    erroresCantidad,
    total
) {

    const registro = {

        fecha:
            new Date().toLocaleString(),

        nota,

        porcentaje,

        aciertos,

        errores:
            erroresCantidad,

        total

    };


    historico.push(
        registro
    );


    localStorage.setItem(
        "historico",
        JSON.stringify(
            historico
        )
    );
}


function cargarHistorico() {

    try {

        historico =
            JSON.parse(
                localStorage.getItem(
                    "historico"
                )
            ) || [];

    }
    catch (error) {

        console.error(
            "Error cargando histórico:",
            error
        );

        historico = [];

    }
}


function mostrarHistorico() {

    if (
        historico.length === 0
    ) {

        alert(
            "Todavía no hay registros."
        );

        return;
    }


    let texto =
        "HISTÓRICO\n\n";


    historico
        .slice()
        .reverse()
        .forEach(
            registro => {

                texto +=
`📅 ${registro.fecha}

Nota: ${registro.nota}/10

Aciertos: ${registro.aciertos}
Errores: ${registro.errores}

Porcentaje:
${registro.porcentaje}%

------------------------

`;

            }
        );


    alert(
        texto
    );
}


/* =====================================
   PARTE 18 - LIMPIAR HISTÓRICO
===================================== */

function limpiarHistorico() {

    const confirmar =
        confirm(
            "¿Desea borrar el histórico y las estadísticas?"
        );


    if (!confirmar) {

        return;

    }


    historico = [];

    estadisticas = {};


    localStorage.removeItem(
        "historico"
    );

    localStorage.removeItem(
        "estadisticasVerbos"
    );


    alert(
        "Información eliminada."
    );
}


/* =====================================
   PARTE 19 - EXPORTAR ERRORES
===================================== */

function exportarErrores() {

    if (
        errores.length === 0
    ) {

        alert(
            "No existen errores."
        );

        return;
    }


    const datos =
        errores.map(
            verbo => ({

                "INFINITIVE":
                    verbo.infinitivo,

                "PAST SIMPLE":
                    verbo.pasado,

                "PAST PARTICIPLE":
                    verbo.participio,

                "MEANING":
                    verbo.meaning

            })
        );


    const ws =
        XLSX.utils.json_to_sheet(
            datos
        );


    const wb =
        XLSX.utils.book_new();


    XLSX.utils.book_append_sheet(
        wb,
        ws,
        "Errores"
    );


    XLSX.writeFile(
        wb,
        "Errores.xlsx"
    );
}


/* =====================================
   PARTE 20 - EXPORTAR RESULTADOS
===================================== */

function exportarResultados() {

    if (
        resultados.length === 0
    ) {

        alert(
            "No existen resultados."
        );

        return;
    }


    const ws =
        XLSX.utils.json_to_sheet(
            resultados
        );


    const wb =
        XLSX.utils.book_new();


    XLSX.utils.book_append_sheet(
        wb,
        ws,
        "Resultados"
    );


    XLSX.writeFile(
        wb,
        "Resultados.xlsx"
    );
}


/* =====================================
   PARTE 21 - ESTADÍSTICAS
===================================== */

function registrarPregunta(
    verbo
) {

    const clave =
        obtenerClaveVerbo(
            verbo
        );


    if (!estadisticas[clave]) {

        estadisticas[clave] =
            crearRegistroEstadistica(
                verbo
            );

    }


    estadisticas[
        clave
    ].vecesPreguntado++;


    guardarEstadisticas();
}


function registrarErrorVerbo(
    verbo
) {

    const clave =
        obtenerClaveVerbo(
            verbo
        );


    if (!estadisticas[clave]) {

        estadisticas[clave] =
            crearRegistroEstadistica(
                verbo
            );

    }


    estadisticas[
        clave
    ].fallos++;


    guardarEstadisticas();
}


function crearRegistroEstadistica(
    verbo
) {

    return {

        infinitivo:
            verbo.infinitivo,

        pasado:
            verbo.pasado,

        participio:
            verbo.participio,

        meaning:
            verbo.meaning,

        fallos: 0,

        vecesPreguntado: 0

    };
}


function obtenerClaveVerbo(
    verbo
) {

    return normalizarTexto(
        verbo.infinitivo
    );
}


function guardarEstadisticas() {

    localStorage.setItem(
        "estadisticasVerbos",
        JSON.stringify(
            estadisticas
        )
    );
}


/* =====================================
   PARTE 22 - MOSTRAR ESTADÍSTICAS
===================================== */

function mostrarEstadisticas() {

    const lista =
        Object.values(
            estadisticas
        ).sort(
            (a, b) =>
                b.fallos -
                a.fallos
        );


    if (
        lista.length === 0
    ) {

        alert(
            "Todavía no existen estadísticas."
        );

        return;
    }


    let html = `

        <div class="panel">

            <h2>
                📊 Estadísticas
            </h2>

            <p>
                Verbos ordenados
                por número de fallos.
            </p>

        </div>

        <div class="tabla-estudio">

            <div class="cabecera-estudio">

                <div>
                    Verbo
                </div>

                <div>
                    Significado
                </div>

                <div>
                    Fallos
                </div>

                <div>
                    % Éxito
                </div>

            </div>

    `;


    lista.forEach(
        verbo => {

            const porcentaje =
                verbo.vecesPreguntado > 0
                    ?
                    Math.round(
                        (
                            (
                                verbo.vecesPreguntado -
                                verbo.fallos
                            ) * 100
                        ) /
                        verbo.vecesPreguntado
                    )
                    :
                    0;


            html += `

                <div class="fila-estudio">

                    <div>

                        ${escapeHTML(
                            verbo.infinitivo
                        )}

                        <br>

                        <small>

                            ${escapeHTML(
                                verbo.pasado
                            )}

                            /

                            ${escapeHTML(
                                verbo.participio
                            )}

                        </small>

                    </div>


                    <div>

                        ${escapeHTML(
                            verbo.meaning
                        )}

                    </div>


                    <div>

                        ${verbo.fallos}

                    </div>


                    <div>

                        ${porcentaje}%

                    </div>

                </div>

            `;

        }
    );


    html += `
        </div>
    `;


    const contenedor =
        document.getElementById(
            "contenido"
        );


    if (contenedor) {

        contenedor.innerHTML =
            html;

    }


    window.scrollTo({

        top: 0,

        behavior: "smooth"

    });
}


/* =====================================
   PARTE 23 - LIMPIAR ESTADÍSTICAS
===================================== */

function limpiarEstadisticas() {

    const confirmar =
        confirm(
            "¿Desea borrar todas las estadísticas?"
        );


    if (!confirmar) {

        return;

    }


    estadisticas = {};


    localStorage.removeItem(
        "estadisticasVerbos"
    );


    alert(
        "Estadísticas eliminadas."
    );
}


/* =====================================
   PARTE 24 - PWA
===================================== */

window.addEventListener(
    "beforeinstallprompt",
    event => {

        event.preventDefault();

        deferredPrompt = event;


        const boton =
            document.getElementById(
                "btnInstalar"
            );


        if (boton) {

            boton.hidden = false;

        }

    }
);


window.addEventListener(
    "appinstalled",
    () => {

        deferredPrompt = null;


        const boton =
            document.getElementById(
                "btnInstalar"
            );


        if (boton) {

            boton.hidden = true;

        }

    }
);


document.addEventListener(
    "click",
    async event => {

        const boton =
            event.target.closest(
                "#btnInstalar"
            );


        if (!boton) {

            return;

        }


        if (!deferredPrompt) {

            alert(
                "La instalación no está disponible en este momento."
            );

            return;

        }


        deferredPrompt.prompt();


        try {

            await deferredPrompt.userChoice;

        }
        catch (error) {

            console.error(
                "Error durante la instalación:",
                error
            );

        }


        deferredPrompt = null;

        boton.hidden = true;

    }
);


/* =====================================
   FUNCIONES AUXILIARES
===================================== */

function normalizarTexto(
    texto
) {

    return String(
        texto ?? ""
    )
        .trim()
        .toLowerCase()
        .replace(
            /\s+/g,
            " "
        );
}


function mezclarArray(
    array
) {

    return [...array].sort(
        () =>
            Math.random() - 0.5
    );
}


function limpiarResultado() {

    const resultado =
        document.getElementById(
            "resultado"
        );


    if (resultado) {

        resultado.innerHTML = "";

    }
}


function escapeHTML(
    texto
) {

    return String(
        texto ?? ""
    )
        .replace(
            /&/g,
            "&amp;"
        )
        .replace(
            /</g,
            "<"
        )
        .replace(
            />/g,
            ">"
        )
        .replace(
            /"/g,
            "&quot;"
        )
        .replace(
            /'/g,
            "&#039;"
        );
}
