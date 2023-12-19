let archivoExcel;
const tabla = document.querySelector('#tabla-prestaciones');
const seccionError = document.querySelector('#seccion-error');

function procesarArchivo(){
    const input = document.querySelector('#archivoInput');
    archivoExcel = input.files[0];
    seccionError.innerHTML = '';
    tabla.innerHTML = '';

    if (archivoExcel){
        interpretarExcel(archivoExcel);
    } else {
        const textoError = document.createElement('p');
        textoError.classList.add('error');
        textoError.textContent = 'No se ha seleccionado ningún archivo';
        seccionError.appendChild(textoError);
    }
}

function interpretarExcel(archivo){
    var reader = new FileReader();
    reader.onload = function(e) {
        var data = new Uint8Array(e.target.result);
        var workbook = XLSX.read(data, { type: 'array' });
        // Itera a través de todas las hojas en el archivo
        workbook.SheetNames.forEach(function(sheetName, index) {
            var worksheet = workbook.Sheets[sheetName];
            var jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            const filaTitulo = document.createElement('tr');
            const encabezado = document.createElement('th');
            encabezado.textContent = sheetName;
            encabezado.setAttribute('colspan', '2');
            encabezado.classList.add('encabezado-h1');
            
            const botonMostrarOcultar = document.createElement('button');
            botonMostrarOcultar.classList.add('boton-toggle');
            botonMostrarOcultar.textContent = 'Mostrar';
            botonMostrarOcultar.addEventListener('click', function() {
                const filaContenido = document.querySelector(`#contenido-${index + 1}`);
                    if (filaContenido.style.display === 'none') {
                        filaContenido.style.display = 'table-row';
                        botonMostrarOcultar.textContent = 'Ocultar';
                    } else {
                        filaContenido.style.display = 'none';
                        botonMostrarOcultar.textContent = 'Mostrar';
                    }
                });

            encabezado.appendChild(botonMostrarOcultar);
            filaTitulo.appendChild(encabezado);

            const filaContenido = document.createElement('tr');
            filaContenido.setAttribute('id', `contenido-${index + 1}`);
            const celdaContenido = document.createElement('td');

            if(sheetName === 'Facturacion' || sheetName === 'Facturación'){
                for(let i = 0; i < jsonData.length; i++){
                    if(jsonData[i][0]){
                        const contenidoParrafo = document.createElement('p');
                        contenidoParrafo.textContent = jsonData[i][0];                            
                        if (jsonData[i][0].endsWith(':')) {
                            const contenidoH3 = document.createElement('h3');
                            contenidoH3.textContent = jsonData[i][0];
                            celdaContenido.appendChild(contenidoH3);    
                        } else {
                            celdaContenido.appendChild(contenidoParrafo);
                        }
                    }   
                }
            }
            else if(!ContenidoB1(worksheet)){  
                celdaContenido.colSpan = 2;                  
                for(let i = 0; i < jsonData.length; i++){
                    if(jsonData[i][0]){
                        const contenidoParrafo = document.createElement('p');
                        contenidoParrafo.textContent = jsonData[i][0];
                        if (i === 0 && jsonData[i].length === 1) {
                            // Si es la celda A1, mostrar el contenido en h3 con estilo especial
                            const contenidoH3 = document.createElement('h3');
                            contenidoH3.textContent = jsonData[i][0];
                            contenidoH3.classList.add('contenidoH3');
                            celdaContenido.appendChild(contenidoH3);
                        } else if (typeof jsonData[i][0] === 'string' &&jsonData[i][0].endsWith(':')) {
                            const contenidoH3 = document.createElement('h3');
                            contenidoH3.textContent = jsonData[i][0];
                            celdaContenido.appendChild(contenidoH3);
                        } else if (typeof jsonData[i][0] === 'string' && jsonData[i][0].endsWith('cluye:')) {
                            const contenidoH4 = document.createElement('h4');
                            contenidoH4.textContent = jsonData[i][0];
                            celdaContenido.appendChild(contenidoH4);
                        } else {
                            celdaContenido.appendChild(contenidoParrafo);
                        }
                    }   
                }
            }
            else {
                for (let i = 0; i < jsonData.length; i++) {
                    const filaCodigos = document.createElement('tr');
        
                    const celdaColumnaA = document.createElement('td');
                    celdaColumnaA.textContent = jsonData[i][0] || '';
        
                    const celdaColumnaB = document.createElement('td');
                    celdaColumnaB.textContent = jsonData[i][1] || '';
        
                    const celdaColumnaC = document.createElement('td');
                    celdaColumnaC.textContent = jsonData[i][2] || '';

                    if(i == 0){
                        celdaColumnaA.style.fontWeight = 'bold';
                        celdaColumnaB.style.fontWeight = 'bold';
                        celdaColumnaB.style.textAlign = 'center';
                    }

                    if(!filaVacia(celdaColumnaA, celdaColumnaB,celdaColumnaC)){
                        filaCodigos.appendChild(celdaColumnaA);
                        filaCodigos.appendChild(celdaColumnaB);
                        filaCodigos.appendChild(celdaColumnaC);
                        celdaContenido.appendChild(filaCodigos);
                    }
                    
                    tabla.appendChild(celdaContenido);
                }
            }

            filaContenido.appendChild(celdaContenido);
            filaContenido.style.display = 'none';
            tabla.appendChild(filaTitulo);
            tabla.appendChild(filaContenido);
        });
    };
    reader.readAsArrayBuffer(archivo);
}

function ContenidoB1(worksheet){
    var cellValue = worksheet['B1'] ? worksheet['B1'].v : '';
    return cellValue.trim();
}

function filaVacia(celdaA, celdaB, celdaC){
    return celdaA.textContent === '' && celdaB.textContent === '' && celdaC.textContent === '';
}