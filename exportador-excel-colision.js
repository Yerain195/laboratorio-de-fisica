// Función para crear gráficos profesionales con Chart.js
function crearGraficoProfesional(tipo, datos, opciones, ancho = 800, alto = 500) {
    return new Promise((resolve) => {
        const canvas = document.createElement('canvas');
        canvas.width = ancho;
        canvas.height = alto;
        const ctx = canvas.getContext('2d');
        
        // Fondo blanco profesional
        ctx.fillStyle = '#FFFFFF';
        ctx.fillRect(0, 0, canvas.width, canvas.height);
        
        // Paleta de colores profesional
        const coloresProfesionales = {
            carrito1: {
                fondo: 'rgba(74, 144, 226, 0.85)',
                borde: 'rgba(74, 144, 226, 1)',
                degradado: ['#4a90e2', '#357abd']
            },
            carrito2: {
                fondo: 'rgba(237, 85, 100, 0.85)',
                borde: 'rgba(237, 85, 100, 1)',
                degradado: ['#ed5564', '#da4453']
            },
            inicial: {
                fondo: 'rgba(76, 175, 80, 0.85)',
                borde: 'rgba(76, 175, 80, 1)'
            },
            final: {
                fondo: 'rgba(156, 39, 176, 0.85)',
                borde: 'rgba(156, 39, 176, 1)'
            }
        };

        // Crear gradientes
        const gradient1 = ctx.createLinearGradient(0, 0, 0, alto);
        gradient1.addColorStop(0, coloresProfesionales.carrito1.degradado[0]);
        gradient1.addColorStop(1, coloresProfesionales.carrito1.degradado[1]);

        const gradient2 = ctx.createLinearGradient(0, 0, 0, alto);
        gradient2.addColorStop(0, coloresProfesionales.carrito2.degradado[0]);
        gradient2.addColorStop(1, coloresProfesionales.carrito2.degradado[1]);

        // Aplicar gradientes a datasets
        if (tipo === 'bar' || tipo === 'line') {
            datos.datasets.forEach((dataset, index) => {
                if (index === 0) {
                    dataset.backgroundColor = gradient1;
                    dataset.borderColor = coloresProfesionales.carrito1.borde;
                } else if (index === 1) {
                    dataset.backgroundColor = gradient2;
                    dataset.borderColor = coloresProfesionales.carrito2.borde;
                }
            });
        }

        // Configuración del gráfico
        new Chart(ctx, {
            type: tipo,
            data: datos,
            options: {
                ...opciones,
                responsive: false,
                maintainAspectRatio: false,
                layout: {
                    padding: {
                        top: 30,
                        right: 30,
                        bottom: 30,
                        left: 30
                    }
                },
                plugins: {
                    legend: {
                        position: 'top',
                        labels: {
                            font: {
                                size: 14,
                                family: 'Segoe UI, Arial, sans-serif',
                                weight: 'bold'
                            },
                            padding: 20,
                            usePointStyle: true,
                            pointStyle: 'circle'
                        }
                    },
                    title: {
                        display: true,
                        font: {
                            size: 18,
                            family: 'Segoe UI, Arial, sans-serif',
                            weight: 'bold'
                        },
                        padding: 25,
                        color: '#2c3e50'
                    },
                    tooltip: {
                        backgroundColor: 'rgba(44, 62, 80, 0.95)',
                        titleFont: {
                            size: 13,
                            family: 'Segoe UI, Arial, sans-serif'
                        },
                        bodyFont: {
                            size: 12,
                            family: 'Segoe UI, Arial, sans-serif'
                        },
                        padding: 12,
                        cornerRadius: 6
                    }
                },
                scales: tipo !== 'doughnut' && tipo !== 'pie' ? {
                    y: {
                        beginAtZero: true,
                        grid: {
                            color: 'rgba(0, 0, 0, 0.08)',
                            drawBorder: false
                        },
                        ticks: {
                            font: {
                                size: 12,
                                family: 'Segoe UI, Arial, sans-serif'
                            },
                            padding: 10
                        },
                        title: {
                            display: true,
                            font: {
                                size: 14,
                                family: 'Segoe UI, Arial, sans-serif',
                                weight: 'bold'
                            },
                            color: '#2c3e50',
                            padding: 12
                        }
                    },
                    x: {
                        grid: {
                            display: false
                        },
                        ticks: {
                            font: {
                                size: 12,
                                family: 'Segoe UI, Arial, sans-serif'
                            },
                            padding: 10
                        },
                        title: {
                            display: true,
                            font: {
                                size: 14,
                                family: 'Segoe UI, Arial, sans-serif',
                                weight: 'bold'
                            },
                            color: '#2c3e50',
                            padding: 12
                        }
                    }
                } : {},
                elements: {
                    bar: {
                        borderRadius: 6,
                        borderWidth: 0
                    },
                    line: {
                        tension: 0.4,
                        borderWidth: 3
                    },
                    point: {
                        radius: 6,
                        hoverRadius: 8,
                        backgroundColor: '#ffffff',
                        borderWidth: 3
                    }
                },
                animation: {
                    duration: 1000,
                    easing: 'easeOutQuart'
                }
            }
        });

        setTimeout(() => {
            const base64 = canvas.toDataURL('image/png', 1.0);
            resolve(base64);
        }, 700);
    });
}

async function generarExcelAvanzado(datos) {
    console.log('🔄 Generando Excel profesional mejorado...', datos);
    
    if (!datos) {
        alert('No hay datos para exportar. Ejecuta la simulación primero.');
        return;
    }

    if (typeof ExcelJS === 'undefined') {
        alert('Error: ExcelJS no está cargado.');
        return;
    }

    if (typeof Chart === 'undefined') {
        alert('Error: Chart.js no está cargado.');
        return;
    }

    try {
        const workbook = new ExcelJS.Workbook();
        workbook.creator = 'Laboratorio Virtual de Física - Universidad';
        workbook.created = new Date();
        workbook.company = 'Departamento de Física';
        
        // ========== HOJA 1: PORTADA Y RESUMEN EJECUTIVO ==========
        const hojaPortada = workbook.addWorksheet('Portada', {
            properties: { tabColor: { argb: 'FF1A237E' } }
        });

        hojaPortada.columns = Array(8).fill({ width: 15 });

        // PORTADA - Título principal
        hojaPortada.mergeCells('B3:G5');
        const tituloPortada = hojaPortada.getCell('B3');
        tituloPortada.value = '⚛️ LABORATORIO VIRTUAL DE FÍSICA';
        tituloPortada.font = { name: 'Calibri', size: 28, bold: true, color: { argb: 'FFFFFFFF' } };
        tituloPortada.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1A237E' } };
        tituloPortada.alignment = { vertical: 'middle', horizontal: 'center' };
        hojaPortada.getRow(3).height = 80;

        // Subtítulo
        hojaPortada.mergeCells('B7:G7');
        const subtitulo = hojaPortada.getCell('B7');
        subtitulo.value = 'ANÁLISIS DE COLISIONES ELÁSTICAS';
        subtitulo.font = { name: 'Calibri', size: 20, bold: true, color: { argb: 'FF1A237E' } };
        subtitulo.alignment = { horizontal: 'center' };
        hojaPortada.getRow(7).height = 30;

        // Información del reporte
        hojaPortada.mergeCells('B9:G9');
        const infoFecha = hojaPortada.getCell('B9');
        infoFecha.value = `📅 Fecha de Experimento: ${new Date().toLocaleDateString('es-ES', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' })}`;
        infoFecha.font = { size: 12, italic: true };
        infoFecha.alignment = { horizontal: 'center' };

        hojaPortada.mergeCells('B10:G10');
        const infoHora = hojaPortada.getCell('B10');
        infoHora.value = `🕐 Hora: ${new Date().toLocaleTimeString('es-ES')}`;
        infoHora.font = { size: 12, italic: true };
        infoHora.alignment = { horizontal: 'center' };

        // Resumen ejecutivo
        hojaPortada.mergeCells('B13:G13');
        const tituloResumen = hojaPortada.getCell('B13');
        tituloResumen.value = '📊 RESUMEN';
        tituloResumen.font = { size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
        tituloResumen.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1976D2' } };
        tituloResumen.alignment = { vertical: 'middle', horizontal: 'center' };
        hojaPortada.getRow(13).height = 25;

        const resumenDatos = [
            ['Parámetro', 'Valor', 'Interpretación'],
            ['Masa Total del Sistema', `${(datos.m1 + datos.m2).toFixed(3)} kg`, 'Suma de masas de ambos carritos'],
            ['Velocidad Relativa Inicial', `${Math.abs(datos.v1 - datos.v2).toFixed(3)} m/s`, 'Velocidad de aproximación'],
            ['Energía Total Inicial', `${datos.ecInicial.toFixed(4)} J`, 'Energía cinética del sistema antes'],
            ['Momento Total Inicial', `${datos.pInicial.toFixed(4)} kg·m/s`, 'Cantidad de movimiento inicial'],
            ['Conservación de Energía', `${((datos.ecFinal/datos.ecInicial)*100).toFixed(2)}%`, 'Porcentaje de energía conservada'],
            ['Conservación de Momento', `${((datos.pFinal/datos.pInicial)*100).toFixed(2)}%`, 'Porcentaje de momento conservado'],
            ['Error de Energía', `${Math.abs(datos.ecInicial - datos.ecFinal).toFixed(6)} J`, 'Diferencia energética'],
            ['Error de Momento', `${Math.abs(datos.pInicial - datos.pFinal).toFixed(6)} kg·m/s`, 'Diferencia de momento']
        ];

        let filaResumen = 14;
        resumenDatos.forEach((fila, idx) => {
            ['B', 'C', 'D', 'E', 'F', 'G'].forEach((col, colIdx) => {
                const celda = hojaPortada.getCell(`${col}${filaResumen}`);
                if (colIdx < 3) {
                    celda.value = fila[colIdx];
                }
            });

            hojaPortada.mergeCells(`B${filaResumen}:C${filaResumen}`);
            hojaPortada.mergeCells(`D${filaResumen}:E${filaResumen}`);
            hojaPortada.mergeCells(`F${filaResumen}:G${filaResumen}`);

            const celdaParam = hojaPortada.getCell(`B${filaResumen}`);
            const celdaValor = hojaPortada.getCell(`D${filaResumen}`);
            const celdaInterp = hojaPortada.getCell(`F${filaResumen}`);

            if (idx === 0) {
                [celdaParam, celdaValor, celdaInterp].forEach(c => {
                    c.font = { bold: true, color: { argb: 'FFFFFFFF' } };
                    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1976D2' } };
                    c.alignment = { horizontal: 'center', vertical: 'middle' };
                    c.border = {
                        top: { style: 'thin' },
                        left: { style: 'thin' },
                        bottom: { style: 'thin' },
                        right: { style: 'thin' }
                    };
                });
            } else {
                celdaParam.font = { bold: true };
                celdaParam.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBBDEFB' } };
                celdaValor.font = { bold: true, size: 11, color: { argb: 'FF0D47A1' } };
                celdaValor.alignment = { horizontal: 'center' };
                celdaInterp.font = { italic: true, size: 10 };
                
                [celdaParam, celdaValor, celdaInterp].forEach(c => {
                    c.alignment = { ...c.alignment, vertical: 'middle' };
                    c.border = {
                        top: { style: 'thin' },
                        left: { style: 'thin' },
                        bottom: { style: 'thin' },
                        right: { style: 'thin' }
                    };
                });
            }
            filaResumen++;
        });

        // ========== HOJA 2: DATOS DETALLADOS ==========
        const hojaDatos = workbook.addWorksheet('Datos Experimentales', {
            properties: { tabColor: { argb: 'FF4472C4' } }
        });

        hojaDatos.columns = [
            { width: 5 },
            { width: 30 },
            { width: 18 },
            { width: 12 },
            { width: 5 }
        ];

        // Título
        hojaDatos.mergeCells('B2:D2');
        const tituloDatos = hojaDatos.getCell('B2');
        tituloDatos.value = '📋 DATOS EXPERIMENTALES COMPLETOS';
        tituloDatos.font = { name: 'Calibri', size: 18, bold: true, color: { argb: 'FFFFFFFF' } };
        tituloDatos.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF203864' } };
        tituloDatos.alignment = { vertical: 'middle', horizontal: 'center' };
        hojaDatos.getRow(2).height = 30;

        // Sección: Condiciones Iniciales
        hojaDatos.getRow(5).height = 25;
        hojaDatos.mergeCells('B5:D5');
        const seccionInicial = hojaDatos.getCell('B5');
        seccionInicial.value = '🔵 CONDICIONES INICIALES';
        seccionInicial.font = { size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
        seccionInicial.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF5B9BD5' } };
        seccionInicial.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };

        const datosIniciales = [
            ['Parámetro', 'Valor', 'Unidad'],
            ['Masa Carrito 1 (m₁)', datos.m1, 'kg'],
            ['Velocidad Inicial Carrito 1 (v₁ᵢ)', datos.v1, 'm/s'],
            ['Momento Inicial Carrito 1', datos.m1 * datos.v1, 'kg·m/s'],
            ['Energía Cinética Inicial Carrito 1', 0.5 * datos.m1 * datos.v1 * datos.v1, 'J'],
            ['Masa Carrito 2 (m₂)', datos.m2, 'kg'],
            ['Velocidad Inicial Carrito 2 (v₂ᵢ)', datos.v2, 'm/s'],
            ['Momento Inicial Carrito 2', datos.m2 * datos.v2, 'kg·m/s'],
            ['Energía Cinética Inicial Carrito 2', 0.5 * datos.m2 * datos.v2 * datos.v2, 'J']
        ];

        let filaActual = 6;
        datosIniciales.forEach((fila, idx) => {
            const celdaB = hojaDatos.getCell(`B${filaActual}`);
            const celdaC = hojaDatos.getCell(`C${filaActual}`);
            const celdaD = hojaDatos.getCell(`D${filaActual}`);
            
            celdaB.value = fila[0];
            celdaC.value = fila[1];
            celdaD.value = fila[2];

            if (idx === 0) {
                [celdaB, celdaC, celdaD].forEach(celda => {
                    celda.font = { bold: true, color: { argb: 'FFFFFFFF' } };
                    celda.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } };
                    celda.alignment = { horizontal: 'center', vertical: 'middle' };
                });
            } else {
                celdaB.font = { bold: true };
                celdaB.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
                celdaC.font = { size: 11 };
                celdaC.alignment = { horizontal: 'center' };
                if (typeof celdaC.value === 'number') {
                    celdaC.numFmt = '0.0000';
                }
                celdaD.font = { italic: true };
                celdaD.alignment = { horizontal: 'center' };
                
                [celdaB, celdaC, celdaD].forEach(celda => {
                    celda.border = {
                        top: { style: 'thin' },
                        left: { style: 'thin' },
                        bottom: { style: 'thin' },
                        right: { style: 'thin' }
                    };
                });
            }
            filaActual++;
        });

        // Sección: Resultados Finales
        filaActual += 2;
        hojaDatos.getRow(filaActual).height = 25;
        hojaDatos.mergeCells(`B${filaActual}:D${filaActual}`);
        const seccionFinal = hojaDatos.getCell(`B${filaActual}`);
        seccionFinal.value = '🎯 RESULTADOS DESPUÉS DE LA COLISIÓN';
        seccionFinal.font = { size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
        seccionFinal.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF70AD47' } };
        seccionFinal.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };

        filaActual++;
        const resultados = [
            ['Parámetro', 'Valor', 'Unidad'],
            ['Velocidad Final Carrito 1 (v₁f)', datos.v1f, 'm/s'],
            ['Cambio de Velocidad Carrito 1 (Δv₁)', datos.v1f - datos.v1, 'm/s'],
            ['Momento Final Carrito 1', datos.m1 * datos.v1f, 'kg·m/s'],
            ['Energía Cinética Final Carrito 1', 0.5 * datos.m1 * datos.v1f * datos.v1f, 'J'],
            ['Velocidad Final Carrito 2 (v₂f)', datos.v2f, 'm/s'],
            ['Cambio de Velocidad Carrito 2 (Δv₂)', datos.v2f - datos.v2, 'm/s'],
            ['Momento Final Carrito 2', datos.m2 * datos.v2f, 'kg·m/s'],
            ['Energía Cinética Final Carrito 2', 0.5 * datos.m2 * datos.v2f * datos.v2f, 'J']
        ];

        resultados.forEach((fila, idx) => {
            const celdaB = hojaDatos.getCell(`B${filaActual}`);
            const celdaC = hojaDatos.getCell(`C${filaActual}`);
            const celdaD = hojaDatos.getCell(`D${filaActual}`);
            
            celdaB.value = fila[0];
            celdaC.value = fila[1];
            celdaD.value = fila[2];

            if (idx === 0) {
                [celdaB, celdaC, celdaD].forEach(celda => {
                    celda.font = { bold: true, color: { argb: 'FFFFFFFF' } };
                    celda.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF70AD47' } };
                    celda.alignment = { horizontal: 'center', vertical: 'middle' };
                });
            } else {
                celdaB.font = { bold: true };
                celdaB.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2EFDA' } };
                celdaC.font = { size: 11, bold: true, color: { argb: 'FF375623' } };
                celdaC.alignment = { horizontal: 'center' };
                if (typeof celdaC.value === 'number') {
                    celdaC.numFmt = '0.0000';
                }
                celdaD.font = { italic: true };
                celdaD.alignment = { horizontal: 'center' };
                
                [celdaB, celdaC, celdaD].forEach(celda => {
                    celda.border = {
                        top: { style: 'thin' },
                        left: { style: 'thin' },
                        bottom: { style: 'thin' },
                        right: { style: 'thin' }
                    };
                });
            }
            filaActual++;
        });

        // Sección: Conservación
        filaActual += 2;
        hojaDatos.getRow(filaActual).height = 25;
        hojaDatos.mergeCells(`B${filaActual}:D${filaActual}`);
        const seccionConservacion = hojaDatos.getCell(`B${filaActual}`);
        seccionConservacion.value = '✓ VERIFICACIÓN DE LEYES DE CONSERVACIÓN';
        seccionConservacion.font = { size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
        seccionConservacion.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC000' } };
        seccionConservacion.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };

        filaActual++;
        const errorEnergia = Math.abs(datos.ecInicial - datos.ecFinal);
        const errorMomento = Math.abs(datos.pInicial - datos.pFinal);
        const verificaciones = [
            ['Ley de Conservación', 'Estado', 'Error Absoluto'],
            ['Conservación de Energía', errorEnergia < 0.01 ? '✓ VERIFICADA' : '✗ NO VERIFICADA', `${errorEnergia.toFixed(8)} J`],
            ['Conservación de Momento', errorMomento < 0.01 ? '✓ VERIFICADA' : '✗ NO VERIFICADA', `${errorMomento.toFixed(8)} kg·m/s`]
        ];

        verificaciones.forEach((fila, idx) => {
            const celdaB = hojaDatos.getCell(`B${filaActual}`);
            const celdaC = hojaDatos.getCell(`C${filaActual}`);
            const celdaD = hojaDatos.getCell(`D${filaActual}`);
            
            celdaB.value = fila[0];
            celdaC.value = fila[1];
            celdaD.value = fila[2];

            if (idx === 0) {
                [celdaB, celdaC, celdaD].forEach(celda => {
                    celda.font = { bold: true, color: { argb: 'FFFFFFFF' } };
                    celda.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC000' } };
                    celda.alignment = { horizontal: 'center', vertical: 'middle' };
                });
            } else {
                celdaB.font = { bold: true };
                celdaB.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
                
                const verificada = fila[1].includes('✓');
                celdaC.font = { size: 12, bold: true, color: { argb: verificada ? 'FF008000' : 'FFFF0000' } };
                celdaC.alignment = { horizontal: 'center' };
                
                celdaD.font = { size: 10 };
                celdaD.alignment = { horizontal: 'center' };
                
                [celdaB, celdaC, celdaD].forEach(celda => {
                    celda.border = {
                        top: { style: 'thin' },
                        left: { style: 'thin' },
                        bottom: { style: 'thin' },
                        right: { style: 'thin' }
                    };
                });
            }
            filaActual++;
        });

        // ========== HOJA 3: GRÁFICOS PROFESIONALES ==========
        const hojaGraficos = workbook.addWorksheet('Análisis Gráfico', {
            properties: { tabColor: { argb: 'FFFF0000' } }
        });

        hojaGraficos.columns = Array(10).fill({ width: 12 });

        // Título
        hojaGraficos.mergeCells('B2:I2');
        const tituloGraficos = hojaGraficos.getCell('B2');
        tituloGraficos.value = '📈 ANÁLISIS GRÁFICO COMPLETO';
        tituloGraficos.font = { name: 'Calibri', size: 20, bold: true, color: { argb: 'FFFFFFFF' } };
        tituloGraficos.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF0000' } };
        tituloGraficos.alignment = { vertical: 'middle', horizontal: 'center' };
        hojaGraficos.getRow(2).height = 35;

        console.log('📊 Generando gráficos profesionales mejorados...');

        // GRÁFICO 1: Comparación de Velocidades
        const graficoVelocidades = await crearGraficoProfesional('bar', {
            labels: ['Velocidad Inicial', 'Velocidad Final', 'Cambio de Velocidad'],
            datasets: [
                {
                    label: 'Carrito 1',
                    data: [datos.v1, datos.v1f, datos.v1f - datos.v1],
                    borderWidth: 2
                },
                {
                    label: 'Carrito 2',
                    data: [datos.v2, datos.v2f, datos.v2f - datos.v2],
                    borderWidth: 2
                }
            ]
        }, {
            plugins: {
                title: {
                    display: true,
                    text: 'COMPARACIÓN DE VELOCIDADES ANTES Y DESPUÉS'
                }
            },
            scales: {
                y: {
                    title: {
                        display: true,
                        text: 'Velocidad (m/s)'
                    }
                },
                x: {
                    title: {
                        display: true,
                        text: 'Estado del Movimiento'
                    }
                }
            }
        }, 800, 500);

        const imagen1 = workbook.addImage({
            base64: graficoVelocidades.split(',')[1],
            extension: 'png',
        });
        hojaGraficos.addImage(imagen1, {
            tl: { col: 1, row: 4 },
            br: { col: 5, row: 24 }
        });

        // GRÁFICO 2: Energía Cinética
        const graficoEnergia = await crearGraficoProfesional('bar', {
            labels: ['Carrito 1', 'Carrito 2', 'Sistema Total'],
            datasets: [
                {
                    label: 'Energía Inicial (J)',
                    data: [
                        0.5 * datos.m1 * datos.v1 * datos.v1,
                        0.5 * datos.m2 * datos.v2 * datos.v2,
                        datos.ecInicial
                    ],
                    backgroundColor: 'rgba(76, 175, 80, 0.85)',
                    borderColor: 'rgba(76, 175, 80, 1)',
                    borderWidth: 2
                },
                {
                    label: 'Energía Final (J)',
                    data: [
                        0.5 * datos.m1 * datos.v1f * datos.v1f,
                        0.5 * datos.m2 * datos.v2f * datos.v2f,
                        datos.ecFinal
                    ],
                    backgroundColor: 'rgba(156, 39, 176, 0.85)',
                    borderColor: 'rgba(156, 39, 176, 1)',
                    borderWidth: 2
                }
            ]
        }, {
            plugins: {
                title: {
                    display: true,
                    text: 'ANÁLISIS DE ENERGÍA CINÉTICA'
                }
            },
            scales: {
                y: {
                    title: {
                        display: true,
                        text: 'Energía (Joules)'
                    }
                },
                x: {
                    title: {
                        display: true,
                        text: 'Componentes del Sistema'
                    }
                }
            }
        }, 800, 500);

        const imagen2 = workbook.addImage({
            base64: graficoEnergia.split(',')[1],
            extension: 'png',
        });
        hojaGraficos.addImage(imagen2, {
            tl: { col: 5, row: 4 },
            br: { col: 9, row: 24 }
        });

        // GRÁFICO 3: Momento Lineal
        const graficoMomento = await crearGraficoProfesional('bar', {
            labels: ['Carrito 1', 'Carrito 2', 'Sistema Total'],
            datasets: [
                {
                    label: 'Momento Inicial (kg·m/s)',
                    data: [
                        datos.m1 * datos.v1,
                        datos.m2 * datos.v2,
                        datos.pInicial
                    ],
                    backgroundColor: 'rgba(255, 152, 0, 0.85)',
                    borderColor: 'rgba(255, 152, 0, 1)',
                    borderWidth: 2
                },
                {
                    label: 'Momento Final (kg·m/s)',
                    data: [
                        datos.m1 * datos.v1f,
                        datos.m2 * datos.v2f,
                        datos.pFinal
                    ],
                    backgroundColor: 'rgba(0, 150, 136, 0.85)',
                    borderColor: 'rgba(0, 150, 136, 1)',
                    borderWidth: 2
                }
            ]
        }, {
            plugins: {
                title: {
                    display: true,
                    text: 'ANÁLISIS DE MOMENTO LINEAL'
                }
            },
            scales: {
                y: {
                    title: {
                        display: true,
                        text: 'Momento (kg·m/s)'
                    }
                },
                x: {
                    title: {
                        display: true,
                        text: 'Componentes del Sistema'
                    }
                }
            }
        }, 800, 500);

        const imagen3 = workbook.addImage({
            base64: graficoMomento.split(',')[1],
            extension: 'png',
        });
        hojaGraficos.addImage(imagen3, {
            tl: { col: 1, row: 26 },
            br: { col: 5, row: 46 }
        });

        // GRÁFICO 4: Evolución Temporal de Velocidades
        const tiempos = Array.from({length: 50}, (_, i) => i * 0.02);
        const tiempoColision = 0.5;
        const velocidades1 = tiempos.map(t => t < tiempoColision ? datos.v1 : datos.v1f);
        const velocidades2 = tiempos.map(t => t < tiempoColision ? datos.v2 : datos.v2f);

        const graficoEvolucion = await crearGraficoProfesional('line', {
            labels: tiempos.map(t => t.toFixed(2)),
            datasets: [
                {
                    label: 'Carrito 1 (m/s)',
                    data: velocidades1,
                    fill: false,
                    borderWidth: 3,
                    pointRadius: 0
                },
                {
                    label: 'Carrito 2 (m/s)',
                    data: velocidades2,
                    fill: false,
                    borderWidth: 3,
                    pointRadius: 0
                }
            ]
        }, {
            plugins: {
                title: {
                    display: true,
                    text: 'EVOLUCIÓN TEMPORAL DE VELOCIDADES'
                }
            },
            scales: {
                y: {
                    title: {
                        display: true,
                        text: 'Velocidad (m/s)'
                    }
                },
                x: {
                    title: {
                        display: true,
                        text: 'Tiempo (segundos)'
                    },
                    ticks: {
                        maxTicksLimit: 10
                    }
                }
            }
        }, 800, 500);

        const imagen4 = workbook.addImage({
            base64: graficoEvolucion.split(',')[1],
            extension: 'png',
        });
        hojaGraficos.addImage(imagen4, {
            tl: { col: 5, row: 26 },
            br: { col: 9, row: 46 }
        });

        // GRÁFICO 5: Distribución de Energía (Pie Chart)
        const energiaC1Inicial = 0.5 * datos.m1 * datos.v1 * datos.v1;
        const energiaC2Inicial = 0.5 * datos.m2 * datos.v2 * datos.v2;
        
        const graficoPieInicial = await crearGraficoProfesional('pie', {
            labels: ['Carrito 1', 'Carrito 2'],
            datasets: [{
                data: [energiaC1Inicial, energiaC2Inicial],
                backgroundColor: [
                    'rgba(74, 144, 226, 0.85)',
                    'rgba(237, 85, 100, 0.85)'
                ],
                borderColor: [
                    'rgba(74, 144, 226, 1)',
                    'rgba(237, 85, 100, 1)'
                ],
                borderWidth: 2
            }]
        }, {
            plugins: {
                title: {
                    display: true,
                    text: 'DISTRIBUCIÓN DE ENERGÍA INICIAL'
                },
                legend: {
                    position: 'bottom'
                }
            }
        }, 600, 500);

        const imagen5 = workbook.addImage({
            base64: graficoPieInicial.split(',')[1],
            extension: 'png',
        });
        hojaGraficos.addImage(imagen5, {
            tl: { col: 1, row: 48 },
            br: { col: 4, row: 66 }
        });

        // GRÁFICO 6: Distribución de Energía Final (Pie Chart)
        const energiaC1Final = 0.5 * datos.m1 * datos.v1f * datos.v1f;
        const energiaC2Final = 0.5 * datos.m2 * datos.v2f * datos.v2f;
        
        const graficoPieFinal = await crearGraficoProfesional('pie', {
            labels: ['Carrito 1', 'Carrito 2'],
            datasets: [{
                data: [energiaC1Final, energiaC2Final],
                backgroundColor: [
                    'rgba(74, 144, 226, 0.85)',
                    'rgba(237, 85, 100, 0.85)'
                ],
                borderColor: [
                    'rgba(74, 144, 226, 1)',
                    'rgba(237, 85, 100, 1)'
                ],
                borderWidth: 2
            }]
        }, {
            plugins: {
                title: {
                    display: true,
                    text: 'DISTRIBUCIÓN DE ENERGÍA FINAL'
                },
                legend: {
                    position: 'bottom'
                }
            }
        }, 600, 500);

        const imagen6 = workbook.addImage({
            base64: graficoPieFinal.split(',')[1],
            extension: 'png',
        });
        hojaGraficos.addImage(imagen6, {
            tl: { col: 4, row: 48 },
            br: { col: 7, row: 66 }
        });

        // GRÁFICO 7: Conservación Porcentual
        const porcentajeEnergiaConservada = (datos.ecFinal / datos.ecInicial) * 100;
        const porcentajeMomentoConservado = (datos.pFinal / datos.pInicial) * 100;

        const graficoConservacion = await crearGraficoProfesional('bar', {
            labels: ['Energía', 'Momento Lineal'],
            datasets: [
                {
                    label: 'Conservación (%)',
                    data: [porcentajeEnergiaConservada, porcentajeMomentoConservado],
                    backgroundColor: [
                        porcentajeEnergiaConservada >= 99 ? 'rgba(76, 175, 80, 0.85)' : 'rgba(255, 152, 0, 0.85)',
                        porcentajeMomentoConservado >= 99 ? 'rgba(76, 175, 80, 0.85)' : 'rgba(255, 152, 0, 0.85)'
                    ],
                    borderColor: [
                        porcentajeEnergiaConservada >= 99 ? 'rgba(76, 175, 80, 1)' : 'rgba(255, 152, 0, 1)',
                        porcentajeMomentoConservado >= 99 ? 'rgba(76, 175, 80, 1)' : 'rgba(255, 152, 0, 1)'
                    ],
                    borderWidth: 2
                },
                {
                    label: 'Meta: 100%',
                    data: [100, 100],
                    backgroundColor: 'rgba(200, 200, 200, 0.3)',
                    borderColor: 'rgba(100, 100, 100, 0.5)',
                    borderWidth: 1,
                    borderDash: [5, 5]
                }
            ]
        }, {
            plugins: {
                title: {
                    display: true,
                    text: 'VERIFICACIÓN DE CONSERVACIÓN (%)'
                }
            },
            scales: {
                y: {
                    title: {
                        display: true,
                        text: 'Porcentaje (%)'
                    },
                    min: 95,
                    max: 105
                },
                x: {
                    title: {
                        display: true,
                        text: 'Magnitudes Físicas'
                    }
                }
            }
        }, 700, 500);

        const imagen7 = workbook.addImage({
            base64: graficoConservacion.split(',')[1],
            extension: 'png',
        });
        hojaGraficos.addImage(imagen7, {
            tl: { col: 7, row: 48 },
            br: { col: 10, row: 66 }
        });

        // ========== HOJA 4: FÓRMULAS Y CÁLCULOS ==========
        const hojaFormulas = workbook.addWorksheet('Fórmulas', {
            properties: { tabColor: { argb: 'FFED7D31' } }
        });

        hojaFormulas.columns = [
            { width: 5 },
            { width: 45 },
            { width: 20 },
            { width: 15 },
            { width: 5 }
        ];

        hojaFormulas.mergeCells('B2:D2');
        const tituloFormulas = hojaFormulas.getCell('B2');
        tituloFormulas.value = '🧮 FÓRMULAS Y CÁLCULOS DETALLADOS';
        tituloFormulas.font = { name: 'Calibri', size: 18, bold: true, color: { argb: 'FFFFFFFF' } };
        tituloFormulas.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFED7D31' } };
        tituloFormulas.alignment = { vertical: 'middle', horizontal: 'center' };
        hojaFormulas.getRow(2).height = 30;

        const formulas = [
            { 
                titulo: 'Velocidad Final Carrito 1', 
                formula: 'v₁f = ((m₁ - m₂) × v₁ᵢ + 2 × m₂ × v₂ᵢ) / (m₁ + m₂)', 
                calculo: `((${datos.m1} - ${datos.m2}) × ${datos.v1} + 2 × ${datos.m2} × ${datos.v2}) / (${datos.m1} + ${datos.m2})`,
                valor: datos.v1f 
            },
            { 
                titulo: 'Velocidad Final Carrito 2', 
                formula: 'v₂f = ((m₂ - m₁) × v₂ᵢ + 2 × m₁ × v₁ᵢ) / (m₁ + m₂)', 
                calculo: `((${datos.m2} - ${datos.m1}) × ${datos.v2} + 2 × ${datos.m1} × ${datos.v1}) / (${datos.m1} + ${datos.m2})`,
                valor: datos.v2f 
            },
            { 
                titulo: 'Energía Cinética Inicial Total', 
                formula: 'ECᵢ = ½ × m₁ × v₁ᵢ² + ½ × m₂ × v₂ᵢ²', 
                calculo: `½ × ${datos.m1} × ${datos.v1}² + ½ × ${datos.m2} × ${datos.v2}²`,
                valor: datos.ecInicial 
            },
            { 
                titulo: 'Energía Cinética Final Total', 
                formula: 'ECf = ½ × m₁ × v₁f² + ½ × m₂ × v₂f²', 
                calculo: `½ × ${datos.m1} × ${datos.v1f.toFixed(4)}² + ½ × ${datos.m2} × ${datos.v2f.toFixed(4)}²`,
                valor: datos.ecFinal 
            },
            { 
                titulo: 'Momento Lineal Inicial', 
                formula: 'Pᵢ = m₁ × v₁ᵢ + m₂ × v₂ᵢ', 
                calculo: `${datos.m1} × ${datos.v1} + ${datos.m2} × ${datos.v2}`,
                valor: datos.pInicial 
            },
            { 
                titulo: 'Momento Lineal Final', 
                formula: 'Pf = m₁ × v₁f + m₂ × v₂f', 
                calculo: `${datos.m1} × ${datos.v1f.toFixed(4)} + ${datos.m2} × ${datos.v2f.toFixed(4)}`,
                valor: datos.pFinal 
            },
            { 
                titulo: 'Cambio de Velocidad Carrito 1', 
                formula: 'Δv₁ = v₁f - v₁ᵢ', 
                calculo: `${datos.v1f.toFixed(4)} - ${datos.v1}`,
                valor: datos.v1f - datos.v1 
            },
            { 
                titulo: 'Cambio de Velocidad Carrito 2', 
                formula: 'Δv₂ = v₂f - v₂ᵢ', 
                calculo: `${datos.v2f.toFixed(4)} - ${datos.v2}`,
                valor: datos.v2f - datos.v2 
            },
            { 
                titulo: 'Error de Energía', 
                formula: '|ECᵢ - ECf|', 
                calculo: `|${datos.ecInicial.toFixed(4)} - ${datos.ecFinal.toFixed(4)}|`,
                valor: Math.abs(datos.ecInicial - datos.ecFinal) 
            },
            { 
                titulo: 'Error de Momento', 
                formula: '|Pᵢ - Pf|', 
                calculo: `|${datos.pInicial.toFixed(4)} - ${datos.pFinal.toFixed(4)}|`,
                valor: Math.abs(datos.pInicial - datos.pFinal) 
            }
        ];

        let filaFormula = 5;
        formulas.forEach((item, index) => {
            // Título de la fórmula
            hojaFormulas.getRow(filaFormula).height = 22;
            hojaFormulas.mergeCells(`B${filaFormula}:C${filaFormula}`);
            const celdaTitulo = hojaFormulas.getCell(`B${filaFormula}`);
            celdaTitulo.value = `${index + 1}. ${item.titulo}`;
            celdaTitulo.font = { bold: true, size: 12, color: { argb: 'FFFFFFFF' } };
            celdaTitulo.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFED7D31' } };
            celdaTitulo.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
            celdaTitulo.border = { 
                top: { style: 'thin' }, 
                left: { style: 'thin' }, 
                bottom: { style: 'thin' }, 
                right: { style: 'thin' } 
            };

            const celdaValor = hojaFormulas.getCell(`D${filaFormula}`);
            celdaValor.value = item.valor;
            celdaValor.numFmt = '0.0000';
            celdaValor.font = { bold: true, size: 12, color: { argb: 'FF974806' } };
            celdaValor.alignment = { horizontal: 'center', vertical: 'middle' };
            celdaValor.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCE4D6' } };
            celdaValor.border = { 
                top: { style: 'thin' }, 
                left: { style: 'thin' }, 
                bottom: { style: 'thin' }, 
                right: { style: 'thin' } 
            };

            filaFormula++;

            // Fórmula general
            hojaFormulas.mergeCells(`B${filaFormula}:D${filaFormula}`);
            const celdaFormula = hojaFormulas.getCell(`B${filaFormula}`);
            celdaFormula.value = `Fórmula: ${item.formula}`;
            celdaFormula.font = { italic: true, size: 11, color: { argb: 'FF333333' } };
            celdaFormula.alignment = { vertical: 'middle', horizontal: 'center' };
            celdaFormula.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFE699' } };
            celdaFormula.border = { 
                top: { style: 'thin' }, 
                left: { style: 'thin' }, 
                bottom: { style: 'thin' }, 
                right: { style: 'thin' } 
            };

            filaFormula++;

            // Cálculo con valores
            hojaFormulas.mergeCells(`B${filaFormula}:D${filaFormula}`);
            const celdaCalculo = hojaFormulas.getCell(`B${filaFormula}`);
            celdaCalculo.value = `Cálculo: ${item.calculo}`;
            celdaCalculo.font = { size: 10, color: { argb: 'FF666666' } };
            celdaCalculo.alignment = { vertical: 'middle', horizontal: 'center' };
            celdaCalculo.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF5E6' } };
            celdaCalculo.border = { 
                top: { style: 'thin' }, 
                left: { style: 'thin' }, 
                bottom: { style: 'thin' }, 
                right: { style: 'thin' } 
            };

            filaFormula += 3;
        });

        // ========== HOJA 5: ANÁLISIS ESTADÍSTICO ==========
        const hojaEstadisticas = workbook.addWorksheet('Análisis Estadístico', {
            properties: { tabColor: { argb: 'FF9C27B0' } }
        });

        hojaEstadisticas.columns = [
            { width: 5 },
            { width: 35 },
            { width: 18 },
            { width: 25 },
            { width: 5 }
        ];

        hojaEstadisticas.mergeCells('B2:D2');
        const tituloEstadisticas = hojaEstadisticas.getCell('B2');
        tituloEstadisticas.value = '📊 ANÁLISIS ESTADÍSTICO Y CONCLUSIONES';
        tituloEstadisticas.font = { name: 'Calibri', size: 18, bold: true, color: { argb: 'FFFFFFFF' } };
        tituloEstadisticas.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF9C27B0' } };
        tituloEstadisticas.alignment = { vertical: 'middle', horizontal: 'center' };
        hojaEstadisticas.getRow(2).height = 30;

        // Tabla de análisis
        const velocidadRelativaInicial = Math.abs(datos.v1 - datos.v2);
        const velocidadRelativaFinal = Math.abs(datos.v1f - datos.v2f);
        const coeficienteRestitucion = velocidadRelativaFinal / velocidadRelativaInicial;
        const masaTotal = datos.m1 + datos.m2;
        const masaReducida = (datos.m1 * datos.m2) / masaTotal;
        const velocidadCentroMasa = (datos.m1 * datos.v1 + datos.m2 * datos.v2) / masaTotal;

        const datosAnalisis = [
            ['Parámetro Analizado', 'Valor', 'Interpretación'],
            ['Velocidad Relativa Inicial', `${velocidadRelativaInicial.toFixed(4)} m/s`, 'Velocidad de aproximación'],
            ['Velocidad Relativa Final', `${velocidadRelativaFinal.toFixed(4)} m/s`, 'Velocidad de separación'],
            ['Coeficiente de Restitución', `${coeficienteRestitucion.toFixed(4)}`, 'e = 1 para colisión elástica'],
            ['Masa Total del Sistema', `${masaTotal.toFixed(4)} kg`, 'Suma de ambas masas'],
            ['Masa Reducida', `${masaReducida.toFixed(4)} kg`, 'Masa efectiva del sistema'],
            ['Velocidad Centro de Masa', `${velocidadCentroMasa.toFixed(4)} m/s`, 'Velocidad constante del CM'],
            ['Razón de Masas (m₁/m₂)', `${(datos.m1/datos.m2).toFixed(4)}`, 'Proporción de masas'],
            ['Energía por unidad de masa', `${(datos.ecInicial/masaTotal).toFixed(4)} J/kg`, 'Energía específica'],
            ['Momento por unidad de masa', `${(datos.pInicial/masaTotal).toFixed(4)} m/s`, 'Momento específico']
        ];

        let filaAnalisis = 5;
        datosAnalisis.forEach((fila, idx) => {
            const celdaB = hojaEstadisticas.getCell(`B${filaAnalisis}`);
            const celdaC = hojaEstadisticas.getCell(`C${filaAnalisis}`);
            const celdaD = hojaEstadisticas.getCell(`D${filaAnalisis}`);
            
            celdaB.value = fila[0];
            celdaC.value = fila[1];
            celdaD.value = fila[2];

            if (idx === 0) {
                [celdaB, celdaC, celdaD].forEach(celda => {
                    celda.font = { bold: true, color: { argb: 'FFFFFFFF' } };
                    celda.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF9C27B0' } };
                    celda.alignment = { horizontal: 'center', vertical: 'middle' };
                });
            } else {
                celdaB.font = { bold: true };
                celdaB.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE1BEE7' } };
                celdaC.font = { size: 11, bold: true, color: { argb: 'FF4A148C' } };
                celdaC.alignment = { horizontal: 'center' };
                celdaD.font = { italic: true, size: 10 };
                
                [celdaB, celdaC, celdaD].forEach(celda => {
                    celda.alignment = { ...celda.alignment, vertical: 'middle' };
                    celda.border = {
                        top: { style: 'thin' },
                        left: { style: 'thin' },
                        bottom: { style: 'thin' },
                        right: { style: 'thin' }
                    };
                });
            }
            filaAnalisis++;
        });

        // Conclusiones
        filaAnalisis += 2;
        hojaEstadisticas.mergeCells(`B${filaAnalisis}:D${filaAnalisis}`);
        const tituloConclusiones = hojaEstadisticas.getCell(`B${filaAnalisis}`);
        tituloConclusiones.value = '✅ CONCLUSIONES DEL EXPERIMENTO';
        tituloConclusiones.font = { size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
        tituloConclusiones.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E7D32' } };
        tituloConclusiones.alignment = { vertical: 'middle', horizontal: 'center' };
        hojaEstadisticas.getRow(filaAnalisis).height = 25;

        const conclusiones = [
            `1. La energía cinética se conservó en un ${porcentajeEnergiaConservada.toFixed(2)}%, validando el modelo de colisión elástica.`,
            `2. El momento lineal se conservó en un ${porcentajeMomentoConservado.toFixed(2)}%, confirmando la ley de conservación del momento.`,
            `3. El coeficiente de restitución calculado es ${coeficienteRestitucion.toFixed(4)}, ${coeficienteRestitucion > 0.99 ? 'muy cercano a 1 (colisión elástica ideal)' : 'indicando pérdidas mínimas de energía'}.`,
            `4. La velocidad del centro de masa se mantuvo constante en ${velocidadCentroMasa.toFixed(4)} m/s durante toda la colisión.`,
            `5. El cambio de velocidad del carrito 1 fue ${(datos.v1f - datos.v1).toFixed(4)} m/s y del carrito 2 fue ${(datos.v2f - datos.v2).toFixed(4)} m/s.`,
            `6. La razón de masas (m₁/m₂ = ${(datos.m1/datos.m2).toFixed(4)}) influyó en la transferencia de momento entre los carritos.`
        ];

        filaAnalisis++;
        conclusiones.forEach((conclusion, idx) => {
            hojaEstadisticas.mergeCells(`B${filaAnalisis}:D${filaAnalisis}`);
            const celdaConclusion = hojaEstadisticas.getCell(`B${filaAnalisis}`);
            celdaConclusion.value = conclusion;
            celdaConclusion.font = { size: 11 };
            celdaConclusion.alignment = { vertical: 'middle', horizontal: 'left', indent: 1, wrapText: true };
            celdaConclusion.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: idx % 2 === 0 ? 'FFC8E6C9' : 'FFFFFF' } };
            celdaConclusion.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
            hojaEstadisticas.getRow(filaAnalisis).height = 35;
            filaAnalisis++;
        });

        // ========== HOJA 6: TABLA COMPARATIVA ==========
        const hojaComparativa = workbook.addWorksheet('Tabla Comparativa', {
            properties: { tabColor: { argb: 'FF00BCD4' } }
        });

        hojaComparativa.columns = [
            { width: 5 },
            { width: 30 },
            { width: 18 },
            { width: 18 },
            { width: 18 },
            { width: 5 }
        ];

        hojaComparativa.mergeCells('B2:E2');
        const tituloComparativa = hojaComparativa.getCell('B2');
        tituloComparativa.value = '📋 TABLA COMPARATIVA COMPLETA';
        tituloComparativa.font = { name: 'Calibri', size: 18, bold: true, color: { argb: 'FFFFFFFF' } };
        tituloComparativa.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00BCD4' } };
        tituloComparativa.alignment = { vertical: 'middle', horizontal: 'center' };
        hojaComparativa.getRow(2).height = 30;

        const tablaComparativa = [
            ['Magnitud Física', 'Carrito 1', 'Carrito 2', 'Sistema Total'],
            ['ANTES DE LA COLISIÓN', '', '', ''],
            ['Masa (kg)', datos.m1, datos.m2, datos.m1 + datos.m2],
            ['Velocidad (m/s)', datos.v1, datos.v2, '-'],
            ['Momento (kg·m/s)', datos.m1 * datos.v1, datos.m2 * datos.v2, datos.pInicial],
            ['Energía Cinética (J)', 0.5 * datos.m1 * datos.v1 * datos.v1, 0.5 * datos.m2 * datos.v2 * datos.v2, datos.ecInicial],
            ['', '', '', ''],
            ['DESPUÉS DE LA COLISIÓN', '', '', ''],
            ['Masa (kg)', datos.m1, datos.m2, datos.m1 + datos.m2],
            ['Velocidad (m/s)', datos.v1f, datos.v2f, '-'],
            ['Momento (kg·m/s)', datos.m1 * datos.v1f, datos.m2 * datos.v2f, datos.pFinal],
            ['Energía Cinética (J)', 0.5 * datos.m1 * datos.v1f * datos.v1f, 0.5 * datos.m2 * datos.v2f * datos.v2f, datos.ecFinal],
            ['', '', '', ''],
            ['CAMBIOS (Δ)', '', '', ''],
            ['Cambio de Velocidad (m/s)', datos.v1f - datos.v1, datos.v2f - datos.v2, '-'],
            ['Cambio de Momento (kg·m/s)', datos.m1 * datos.v1f - datos.m1 * datos.v1, datos.m2 * datos.v2f - datos.m2 * datos.v2, datos.pFinal - datos.pInicial],
            ['Cambio de Energía (J)', (0.5 * datos.m1 * datos.v1f * datos.v1f) - (0.5 * datos.m1 * datos.v1 * datos.v1), (0.5 * datos.m2 * datos.v2f * datos.v2f) - (0.5 * datos.m2 * datos.v2 * datos.v2), datos.ecFinal - datos.ecInicial]
        ];

        let filaComparativa = 5;
        tablaComparativa.forEach((fila, idx) => {
            const celdaB = hojaComparativa.getCell(`B${filaComparativa}`);
            const celdaC = hojaComparativa.getCell(`C${filaComparativa}`);
            const celdaD = hojaComparativa.getCell(`D${filaComparativa}`);
            const celdaE = hojaComparativa.getCell(`E${filaComparativa}`);
            
            celdaB.value = fila[0];
            celdaC.value = fila[1];
            celdaD.value = fila[2];
            celdaE.value = fila[3];

            // Filas de encabezado principal
            if (idx === 0) {
                [celdaB, celdaC, celdaD, celdaE].forEach(celda => {
                    celda.font = { bold: true, size: 12, color: { argb: 'FFFFFFFF' } };
                    celda.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00BCD4' } };
                    celda.alignment = { horizontal: 'center', vertical: 'middle' };
                    celda.border = {
                        top: { style: 'medium' },
                        left: { style: 'medium' },
                        bottom: { style: 'medium' },
                        right: { style: 'medium' }
                    };
                });
                hojaComparativa.getRow(filaComparativa).height = 25;
            }
            // Filas de sección (ANTES, DESPUÉS, CAMBIOS)
            else if ([1, 7, 13].includes(idx)) {
                hojaComparativa.mergeCells(`B${filaComparativa}:E${filaComparativa}`);
                celdaB.font = { bold: true, size: 13, color: { argb: 'FFFFFFFF' } };
                celdaB.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: idx === 1 ? 'FF0288D1' : idx === 7 ? 'FF388E3C' : 'FFFF6F00' } };
                celdaB.alignment = { horizontal: 'center', vertical: 'middle' };
                celdaB.border = {
                    top: { style: 'medium' },
                    left: { style: 'medium' },
                    bottom: { style: 'medium' },
                    right: { style: 'medium' }
                };
                hojaComparativa.getRow(filaComparativa).height = 22;
            }
            // Filas vacías
            else if ([6, 12].includes(idx)) {
                // Dejar vacío
            }
            // Filas de datos
            else {
                celdaB.font = { bold: true, size: 11 };
                celdaB.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFB2EBF2' } };
                celdaB.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
                
                [celdaC, celdaD, celdaE].forEach(celda => {
                    celda.font = { size: 11 };
                    celda.alignment = { horizontal: 'center', vertical: 'middle' };
                    if (typeof celda.value === 'number') {
                        celda.numFmt = '0.0000';
                        celda.font = { ...celda.font, bold: true, color: { argb: idx > 12 ? 'FFFF6F00' : 'FF0277BD' } };
                    }
                });
                
                [celdaB, celdaC, celdaD, celdaE].forEach(celda => {
                    celda.border = {
                        top: { style: 'thin' },
                        left: { style: 'thin' },
                        bottom: { style: 'thin' },
                        right: { style: 'thin' }
                    };
                });
            }
            filaComparativa++;
        });

        // ========== HOJA 7: RECOMENDACIONES Y NOTAS ==========
        const hojaNotas = workbook.addWorksheet('Notas y Recomendaciones', {
            properties: { tabColor: { argb: 'FFFF9800' } }
        });

        hojaNotas.columns = [
            { width: 5 },
            { width: 80 },
            { width: 5 }
        ];

        hojaNotas.mergeCells('B2:B2');
        const tituloNotas = hojaNotas.getCell('B2');
        tituloNotas.value = '📝 NOTAS TÉCNICAS Y RECOMENDACIONES';
        tituloNotas.font = { name: 'Calibri', size: 18, bold: true, color: { argb: 'FFFFFFFF' } };
        tituloNotas.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF9800' } };
        tituloNotas.alignment = { vertical: 'middle', horizontal: 'center' };
        hojaNotas.getRow(2).height = 30;

        const notasTecnicas = [
            {
                titulo: '🔬 SOBRE EL EXPERIMENTO',
                contenido: [
                    'Este experimento simula una colisión elástica ideal entre dos carritos en un sistema sin fricción.',
                    'En la realidad, siempre existen pérdidas de energía por fricción, deformaciones y sonido.',
                    'Los valores obtenidos representan el comportamiento teórico ideal del sistema.'
                ]
            },
            {
                titulo: '⚠️ FUENTES DE ERROR',
                contenido: [
                    '• Fricción en el riel o superficie de contacto',
                    '• Resistencia del aire (despreciable a bajas velocidades)',
                    '• Deformaciones en los carritos durante el impacto',
                    '• Errores de medición en masas y velocidades',
                    '• Imprecisiones en los instrumentos de medida'
                ]
            },
            {
                titulo: '✅ CRITERIOS DE VALIDACIÓN',
                contenido: [
                    `• Conservación de energía: ${porcentajeEnergiaConservada.toFixed(2)}% ${porcentajeEnergiaConservada >= 99 ? '✓ EXCELENTE' : porcentajeEnergiaConservada >= 95 ? '✓ BUENO' : '⚠ REVISAR'}`,
                    `• Conservación de momento: ${porcentajeMomentoConservado.toFixed(2)}% ${porcentajeMomentoConservado >= 99 ? '✓ EXCELENTE' : porcentajeMomentoConservado >= 95 ? '✓ BUENO' : '⚠ REVISAR'}`,
                    `• Coeficiente de restitución: ${coeficienteRestitucion.toFixed(4)} ${coeficienteRestitucion >= 0.99 ? '✓ ELÁSTICA' : coeficienteRestitucion >= 0.8 ? '~ CASI ELÁSTICA' : '⚠ INELÁSTICA'}`,
                    '• Error de energía < 0.01 J para considerarse despreciable',
                    '• Error de momento < 0.01 kg·m/s para considerarse despreciable'
                ]
            },
            {
                titulo: '🎓 CONCEPTOS IMPORTANTES',
                contenido: [
                    '• Colisión Elástica: Se conservan tanto el momento como la energía cinética',
                    '• Momento Lineal (p): Producto de masa por velocidad (p = mv)',
                    '• Energía Cinética (Ec): Energía asociada al movimiento (Ec = ½mv²)',
                    '• Centro de Masa: Punto donde se concentra toda la masa del sistema',
                    '• Coeficiente de Restitución: Medida de elasticidad de la colisión (e = 1 para elástica)',
                    '• Masa Reducida: Masa efectiva en problemas de dos cuerpos (μ = m₁m₂/(m₁+m₂))'
                ]
            },
            {
                titulo: '📚 APLICACIONES PRÁCTICAS',
                contenido: [
                    '• Diseño de sistemas de seguridad en vehículos (airbags, zonas de deformación)',
                    '• Análisis de colisiones en deportes (billar, bowling, hockey)',
                    '• Física de partículas (colisiones en aceleradores)',
                    '• Dinámica de asteroides y planetas',
                    '• Diseño de amortiguadores y sistemas de suspensión',
                    '• Juegos y simulaciones físicas en videojuegos'
                ]
            },
            {
                titulo: '🔄 RECOMENDACIONES PARA MEJORAR',
                contenido: [
                    '1. Repetir el experimento varias veces y calcular promedios',
                    '2. Utilizar diferentes combinaciones de masas para observar patrones',
                    '3. Variar las velocidades iniciales sistemáticamente',
                    '4. Comparar con colisiones inelásticas (objetos que se quedan pegados)',
                    '5. Documentar las condiciones experimentales (temperatura, superficie, etc.)',
                    '6. Calibrar los instrumentos de medición antes de cada serie de experimentos'
                ]
            },
            {
                titulo: '💡 REFLEXIONES FINALES',
                contenido: [
                    `Para este experimento específico con m₁=${datos.m1} kg, v₁=${datos.v1} m/s, m₂=${datos.m2} kg, v₂=${datos.v2} m/s:`,
                    `• El carrito ${Math.abs(datos.v1f - datos.v1) > Math.abs(datos.v2f - datos.v2) ? '1' : '2'} experimentó el mayor cambio de velocidad`,
                    `• La energía se distribuyó ${energiaC1Final > energiaC2Final ? 'mayormente en el carrito 1' : 'mayormente en el carrito 2'} después de la colisión`,
                    `• El sistema ${porcentajeEnergiaConservada >= 99 ? 'se comportó de manera casi ideal' : 'presentó pérdidas energéticas medibles'}`,
                    '• Los resultados son consistentes con las predicciones teóricas de la mecánica clásica'
                ]
            }
        ];

        let filaNotas = 5;
        notasTecnicas.forEach((seccion, secIdx) => {
            // Título de sección
            const celdaTitulo = hojaNotas.getCell(`B${filaNotas}`);
            celdaTitulo.value = seccion.titulo;
            celdaTitulo.font = { bold: true, size: 13, color: { argb: 'FFFFFFFF' } };
            celdaTitulo.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF9800' } };
            celdaTitulo.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
            celdaTitulo.border = {
                top: { style: 'medium' },
                left: { style: 'medium' },
                bottom: { style: 'medium' },
                right: { style: 'medium' }
            };
            hojaNotas.getRow(filaNotas).height = 25;
            filaNotas++;

            // Contenido
            seccion.contenido.forEach((linea, lineIdx) => {
                const celdaContenido = hojaNotas.getCell(`B${filaNotas}`);
                celdaContenido.value = linea;
                celdaContenido.font = { size: 11 };
                celdaContenido.alignment = { vertical: 'top', horizontal: 'left', indent: 2, wrapText: true };
                celdaContenido.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: lineIdx % 2 === 0 ? 'FFFFE0B2' : 'FFFFFFFF' } };
                celdaContenido.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
                hojaNotas.getRow(filaNotas).height = linea.length > 80 ? 40 : 25;
                filaNotas++;
            });

            filaNotas += 2; // Espacio entre secciones
        });

        // Nota final
        filaNotas += 2;
        hojaNotas.mergeCells(`B${filaNotas}:B${filaNotas}`);
        const notaFinal = hojaNotas.getCell(`B${filaNotas}`);
        notaFinal.value = `📅 Reporte generado el ${new Date().toLocaleString('es-ES', { 
            weekday: 'long', 
            year: 'numeric', 
            month: 'long', 
            day: 'numeric',
            hour: '2-digit',
            minute: '2-digit',
            second: '2-digit'
        })} | Laboratorio Virtual de Física © 2025`;
        notaFinal.font = { italic: true, size: 10, color: { argb: 'FF666666' } };
        notaFinal.alignment = { horizontal: 'center', vertical: 'middle' };
        notaFinal.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF5F5F5' } };
        notaFinal.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
        hojaNotas.getRow(filaNotas).height = 30;

        // ========== GENERAR Y DESCARGAR ARCHIVO ==========
        console.log('📦 Generando archivo Excel profesional mejorado...');
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { 
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
        });
        
        const url = window.URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        const nombreArchivo = `Colision_Elastica_Completo_${new Date().toISOString().slice(0,10)}_${new Date().getHours()}h${new Date().getMinutes()}m.xlsx`;
        link.download = nombreArchivo;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        window.URL.revokeObjectURL(url);
        
        console.log('✅ Excel profesional mejorado generado exitosamente');
        console.log(`📊 Archivo: ${nombreArchivo}`);
        console.log(`📄 Hojas: 7 (Portada, Datos, Gráficos, Fórmulas, Estadísticas, Comparativa, Notas)`);
        console.log(`📈 Gráficos: 7 gráficos profesionales con dimensiones optimizadas`);
        
    } catch (error) {
        console.error('❌ Error al generar Excel:', error);
        alert('Error al generar el archivo Excel: ' + error.message);
    }
}

// Exportar función global
window.generarExcelAvanzado = generarExcelAvanzado;

console.log('📊 Exportador Excel profesional mejorado cargado correctamente');
console.log('✨ Características: 7 hojas, 7 gráficos, análisis completo, recomendaciones');