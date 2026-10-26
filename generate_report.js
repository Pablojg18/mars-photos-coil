const fs = require('fs');
const path = require('path');
const {
    Document,
    Packer,
    Paragraph,
    TextRun,
    HeadingLevel,
    AlignmentType,
    ImageRun,
    Table,
    TableRow,
    TableCell,
    BorderStyle,
    WidthType,
    convertInchesToTwip,
    ShadingType
} = require('docx');

function readImageBuffer(relPath) {
    const fullPath = path.resolve(__dirname, relPath);
    return fs.readFileSync(fullPath);
}

const img1 = readImageBuffer('screenshots/01_mars_photos_grid.png');
const img2 = readImageBuffer('screenshots/02_mars_photos_scrolled.png');
const img3 = readImageBuffer('screenshots/03_mars_photos_error.png');
const img4 = readImageBuffer('screenshots/04_mars_photos_recovered.png');

const primaryColor = "1E3A8A";   // Deep Indigo/Navy
const secondaryColor = "2563EB"; // Vibrant Blue
const textDark = "1F2937";       // Gray 800
const mutedText = "4B5563";      // Gray 600
const bgHighlight = "F3F4F6";    // Gray 100

const doc = new Document({
    styles: {
        default: {
            document: {
                run: {
                    font: "Arial",
                    size: 22, // 11pt
                    color: textDark,
                },
                paragraph: {
                    spacing: { line: 280, before: 100, after: 100 },
                },
            },
        },
    },
    sections: [
        {
            properties: {
                page: {
                    margin: {
                        top: convertInchesToTwip(1),
                        right: convertInchesToTwip(1),
                        bottom: convertInchesToTwip(1),
                        left: convertInchesToTwip(1),
                    },
                },
            },
            children: [
                // PORTADA / ENCABEZADO
                new Paragraph({
                    text: "INSTITUTO TECNOLÓGICO SUPERIOR DEL SUR DE GUANAJUATO",
                    heading: HeadingLevel.HEADING_2,
                    alignment: AlignmentType.CENTER,
                    spacing: { before: 200, after: 100 },
                    run: { color: primaryColor, bold: true, size: 24 }
                }),
                new Paragraph({
                    text: "INGENIERÍA EN SISTEMAS COMPUTACIONALES / INFORMÁTICA",
                    alignment: AlignmentType.CENTER,
                    spacing: { after: 200 },
                    run: { color: mutedText, bold: true, size: 20 }
                }),
                new Paragraph({
                    text: "PROGRAMACIÓN MÓVIL I",
                    alignment: AlignmentType.CENTER,
                    spacing: { after: 300 },
                    run: { color: secondaryColor, bold: true, size: 28 }
                }),
                new Paragraph({
                    text: "REPORTE DE PRÁCTICA DE LABORATORIO",
                    heading: HeadingLevel.HEADING_1,
                    alignment: AlignmentType.CENTER,
                    spacing: { before: 200, after: 200 },
                    run: { color: primaryColor, bold: true, size: 36 }
                }),
                new Paragraph({
                    text: "Implementación de Coil para Carga Asíncrona de Imágenes y Cuadrícula Adaptativa en Jetpack Compose (Mars Photos App)",
                    alignment: AlignmentType.CENTER,
                    spacing: { after: 400 },
                    run: { italic: true, size: 24, color: mutedText }
                }),

                // TABLA DE METADATOS
                new Table({
                    width: { size: 100, type: WidthType.PERCENTAGE },
                    rows: [
                        new TableRow({
                            children: [
                                new TableCell({
                                    width: { size: 30, type: WidthType.PERCENTAGE },
                                    shading: { type: ShadingType.CLEAR, fill: "E0E7FF" },
                                    children: [new Paragraph({ children: [new TextRun({ text: "Estudiante:", bold: true, color: primaryColor })] })],
                                }),
                                new TableCell({
                                    width: { size: 70, type: WidthType.PERCENTAGE },
                                    children: [new Paragraph({ children: [new TextRun("Pablo Juárez (No. Control: s23120166)")] })],
                                }),
                            ],
                        }),
                        new TableRow({
                            children: [
                                new TableCell({
                                    width: { size: 30, type: WidthType.PERCENTAGE },
                                    shading: { type: ShadingType.CLEAR, fill: "E0E7FF" },
                                    children: [new Paragraph({ children: [new TextRun({ text: "Repositorio GitHub:", bold: true, color: primaryColor })] })],
                                }),
                                new TableCell({
                                    width: { size: 70, type: WidthType.PERCENTAGE },
                                    children: [new Paragraph({ children: [new TextRun({ text: "https://github.com/Pablojg18/mars-photos-coil", bold: true, color: secondaryColor })] })],
                                }),
                            ],
                        }),
                        new TableRow({
                            children: [
                                new TableCell({
                                    width: { size: 30, type: WidthType.PERCENTAGE },
                                    shading: { type: ShadingType.CLEAR, fill: "E0E7FF" },
                                    children: [new Paragraph({ children: [new TextRun({ text: "Tecnologías:", bold: true, color: primaryColor })] })],
                                }),
                                new TableCell({
                                    width: { size: 70, type: WidthType.PERCENTAGE },
                                    children: [new Paragraph({ children: [new TextRun("Kotlin 1.9, Jetpack Compose Material 3, Coil Compose 2.4.0, Retrofit 2, Coroutines, JUnit")] })],
                                }),
                            ],
                        }),
                        new TableRow({
                            children: [
                                new TableCell({
                                    width: { size: 30, type: WidthType.PERCENTAGE },
                                    shading: { type: ShadingType.CLEAR, fill: "E0E7FF" },
                                    children: [new Paragraph({ children: [new TextRun({ text: "Fecha de Entrega:", bold: true, color: primaryColor })] })],
                                }),
                                new TableCell({
                                    width: { size: 70, type: WidthType.PERCENTAGE },
                                    children: [new Paragraph({ children: [new TextRun("Octubre 2026")] })],
                                }),
                            ],
                        }),
                    ],
                }),

                new Paragraph({ text: "", spacing: { after: 300 } }),

                // SECCIÓN 1: INTRODUCCIÓN Y OBJETIVOS
                new Paragraph({
                    text: "1. Introducción y Objetivos",
                    heading: HeadingLevel.HEADING_2,
                    spacing: { before: 300, after: 150 },
                    run: { color: primaryColor, bold: true, size: 28 }
                }),
                new Paragraph({
                    children: [
                        new TextRun("En el desarrollo de aplicaciones móviles contemporáneas, el consumo de servicios web y la presentación interactiva de contenido multimedia constituyen pilares fundamentales. La aplicación "),
                        new TextRun({ text: "Mars Photos", bold: true }),
                        new TextRun(" es un proyecto formativo de Android que consume una API REST pública para obtener fotografías reales capturadas por vehículos exploradores en Marte."),
                    ],
                    spacing: { after: 150 },
                }),
                new Paragraph({
                    children: [
                        new TextRun("El estado inicial (starter) del proyecto únicamente se limitaba a consultar la API mediante Retrofit y desplegar en pantalla un mensaje textual con el recuento total de registros recibidos. El propósito de esta práctica consistió en transformar dicha interfaz en una experiencia de usuario completa y atractiva, integrando "),
                        new TextRun({ text: "Coil (Coroutine Image Loader)", bold: true }),
                        new TextRun(" para descargar y almacenar en caché imágenes remotas de manera eficiente, así como diseñar una cuadrícula adaptativa con "),
                        new TextRun({ text: "LazyVerticalGrid", bold: true }),
                        new TextRun(" de Jetpack Compose."),
                    ],
                    spacing: { after: 150 },
                }),
                new Paragraph({
                    children: [new TextRun({ text: "Objetivos de la práctica:", bold: true })],
                    spacing: { after: 100 },
                }),
                new Paragraph({
                    children: [
                        new TextRun("• Modificar el estado de la interfaz de usuario en el ViewModel ("),
                        new TextRun({ text: "MarsUiState", bold: true }),
                        new TextRun(") para exponer una lista de objetos "),
                        new TextRun({ text: "MarsPhoto", italic: true }),
                        new TextRun(" en lugar de una simple cadena de texto."),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun("• Integrar el componente composable "),
                        new TextRun({ text: "AsyncImage", bold: true }),
                        new TextRun(" de Coil, gestionando placeholders de carga, transiciones suaves (crossfade) e imágenes de error en caso de fallo de red."),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun("• Construir una cuadrícula de imágenes interactiva y fluida mediante "),
                        new TextRun({ text: "LazyVerticalGrid", bold: true }),
                        new TextRun(" con celdas de tamaño adaptativo (150 dp) y diseño de tarjetas con elevación."),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun("• Diseñar una pantalla de error con botón de reintento funcional ("),
                        new TextRun({ text: "Retry Action", italic: true }),
                        new TextRun(") que permita restablecer el flujo de la aplicación sin reiniciar el proceso."),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun("• Validar la arquitectura y fiabilidad del sistema mediante pruebas unitarias automatizadas con fakes y corrutinas."),
                    ],
                    spacing: { after: 250 },
                }),

                // SECCIÓN 2: ARQUITECTURA Y MODIFICACIONES
                new Paragraph({
                    text: "2. Arquitectura de Software y Cambios Técnicos",
                    heading: HeadingLevel.HEADING_2,
                    spacing: { before: 200, after: 150 },
                    run: { color: primaryColor, bold: true, size: 28 }
                }),
                new Paragraph({
                    children: [
                        new TextRun("El proyecto se estructura bajo el patrón arquitectónico oficial recomendado por Google: "),
                        new TextRun({ text: "MVVM (Model - View - ViewModel)", bold: true }),
                        new TextRun(" con flujo unidireccional de datos (UDF) y separación estricta de responsabilidades:"),
                    ],
                    spacing: { after: 150 },
                }),
                new Paragraph({
                    children: [
                        new TextRun({ text: "a) Capa de Estado y ViewModel (MarsViewModel.kt): ", bold: true }),
                        new TextRun("Se actualizó la interfaz sellada "),
                        new TextRun({ text: "MarsUiState", bold: true }),
                        new TextRun(". La variante "),
                        new TextRun({ text: "Success", bold: true }),
                        new TextRun(" ahora encapsula "),
                        new TextRun({ text: "val photos: List<MarsPhoto>", bold: true }),
                        new TextRun(". En la función "),
                        new TextRun({ text: "getMarsPhotos()", bold: true }),
                        new TextRun(", al completarse la llamada asíncrona a "),
                        new TextRun({ text: "marsPhotosRepository.getMarsPhotos()", italic: true }),
                        new TextRun(", el resultado alimenta de inmediato dicho estado."),
                    ],
                    spacing: { after: 100 },
                }),
                new Paragraph({
                    children: [
                        new TextRun({ text: "b) Capa de UI - Tarjeta Individual (MarsPhotoCard): ", bold: true }),
                        new TextRun("Se encapsuló la vista de cada foto dentro de un composable "),
                        new TextRun({ text: "Card", bold: true }),
                        new TextRun(" con elevación de 8 dp. En su interior se utiliza "),
                        new TextRun({ text: "AsyncImage", bold: true }),
                        new TextRun(" configurado con "),
                        new TextRun({ text: "ImageRequest.Builder", italic: true }),
                        new TextRun(", activando animación "),
                        new TextRun({ text: "crossfade(true)", italic: true }),
                        new TextRun(", recurso temporal de espera ("),
                        new TextRun({ text: "loading_img", italic: true }),
                        new TextRun(") e imagen de ruptura ante errores ("),
                        new TextRun({ text: "ic_broken_image", italic: true }),
                        new TextRun("), aplicando además "),
                        new TextRun({ text: "ContentScale.Crop", italic: true }),
                        new TextRun(" para conservar una estética uniforme."),
                    ],
                    spacing: { after: 100 },
                }),
                new Paragraph({
                    children: [
                        new TextRun({ text: "c) Capa de UI - Cuadrícula Adaptativa (PhotosGridScreen): ", bold: true }),
                        new TextRun("Se empleó "),
                        new TextRun({ text: "LazyVerticalGrid(columns = GridCells.Adaptive(150.dp))", bold: true }),
                        new TextRun(", lo que garantiza que la interfaz se ajuste fluidamente a pantallas de cualquier densidad y orientación (vertical u horizontal), reciclando los componentes que salen del viewport para un consumo mínimo de memoria RAM."),
                    ],
                    spacing: { after: 100 },
                }),
                new Paragraph({
                    children: [
                        new TextRun({ text: "d) Manejo de Reintentos (ErrorScreen & MarsPhotosApp): ", bold: true }),
                        new TextRun("Se conectó el botón de reintento en "),
                        new TextRun({ text: "ErrorScreen", bold: true }),
                        new TextRun(" directamente a la lambda "),
                        new TextRun({ text: "marsViewModel::getMarsPhotos", bold: true }),
                        new TextRun(", permitiendo al usuario solicitar nuevamente los datos cuando se restablezca la conectividad."),
                    ],
                    spacing: { after: 250 },
                }),

                // SECCIÓN 3: RESULTADOS Y CAPTURAS
                new Paragraph({
                    text: "3. Resultados Visuales y Evidencias de Ejecución",
                    heading: HeadingLevel.HEADING_2,
                    spacing: { before: 200, after: 150 },
                    run: { color: primaryColor, bold: true, size: 28 }
                }),
                new Paragraph({
                    children: [
                        new TextRun("La aplicación fue instalada y ejecutada en un emulador Android (Android 14 / API 34+). A continuación, se detallan los diferentes estados y comportamientos verificados durante las pruebas dinámicas:"),
                    ],
                    spacing: { after: 200 },
                }),

                // SUBSECCIÓN 3.1: CUADRÍCULA PRINCIPAL
                new Paragraph({
                    text: "3.1. Cuadrícula Principal de Fotografías de Marte",
                    heading: HeadingLevel.HEADING_3,
                    spacing: { before: 150, after: 100 },
                    run: { color: secondaryColor, bold: true, size: 24 }
                }),
                new Paragraph({
                    children: [
                        new TextRun("Al abrir la aplicación, el estado de carga (LoadingScreen con animación de rotación) se muestra brevemente mientras Retrofit descarga el arreglo JSON de la API. Inmediatamente después, Coil procesa en paralelo las URLs remotas, renderizándolas dentro de la cuadrícula adaptativa."),
                    ],
                    spacing: { after: 150 },
                }),
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                        new ImageRun({
                            data: img1,
                            transformation: { width: 260, height: 550 },
                        }),
                    ],
                    spacing: { after: 100 },
                }),
                new Paragraph({
                    text: "Figura 1: Pantalla principal desplegando la cuadrícula de fotografías con Coil y Jetpack Compose.",
                    alignment: AlignmentType.CENTER,
                    spacing: { after: 250 },
                    run: { italic: true, size: 18, color: mutedText }
                }),

                // SUBSECCIÓN 3.2: SCROLL Y CARGA PEREZOSA
                new Paragraph({
                    text: "3.2. Desplazamiento y Carga Perezosa (Lazy Loading)",
                    heading: HeadingLevel.HEADING_3,
                    spacing: { before: 150, after: 100 },
                    run: { color: secondaryColor, bold: true, size: 24 }
                }),
                new Paragraph({
                    children: [
                        new TextRun("Al realizar scroll a lo largo de la cuadrícula, "),
                        new TextRun({ text: "LazyVerticalGrid", bold: true }),
                        new TextRun(" compone únicamente los elementos visibles en el viewport y recicla los que salen de la vista. Coil gestiona una memoria caché de dos niveles (L1 en memoria y L2 en disco), eliminando parpadeos y evitando descargas de red innecesarias para imágenes ya visualizadas."),
                    ],
                    spacing: { after: 150 },
                }),
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                        new ImageRun({
                            data: img2,
                            transformation: { width: 260, height: 550 },
                        }),
                    ],
                    spacing: { after: 100 },
                }),
                new Paragraph({
                    text: "Figura 2: Vista tras desplazamiento (scroll vertical), demostrando carga progresiva y fluidez de fotogramas.",
                    alignment: AlignmentType.CENTER,
                    spacing: { after: 250 },
                    run: { italic: true, size: 18, color: mutedText }
                }),

                // SUBSECCIÓN 3.3: MANEJO DE ERROR Y REINTENTO
                new Paragraph({
                    text: "3.3. Simulación de Pérdida de Conexión y Estado de Error",
                    heading: HeadingLevel.HEADING_3,
                    spacing: { before: 150, after: 100 },
                    run: { color: secondaryColor, bold: true, size: 24 }
                }),
                new Paragraph({
                    children: [
                        new TextRun("Para corroborar la robustez ante fallas de red, se deshabilitaron las interfaces Wi-Fi y datos móviles mediante el comando "),
                        new TextRun({ text: "adb shell svc wifi disable", italic: true }),
                        new TextRun(". El repositorio capturó la excepción "),
                        new TextRun({ text: "IOException", italic: true }),
                        new TextRun(" y el ViewModel actualizó el estado a "),
                        new TextRun({ text: "MarsUiState.Error", bold: true }),
                        new TextRun(". Como se aprecia en la Figura 3, se despliega un gráfico de desconexión, el mensaje 'Failed to load' y un botón 'Retry'."),
                    ],
                    spacing: { after: 150 },
                }),
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                        new ImageRun({
                            data: img3,
                            transformation: { width: 260, height: 550 },
                        }),
                    ],
                    spacing: { after: 100 },
                }),
                new Paragraph({
                    text: "Figura 3: Pantalla de error cuando no hay conexión a Internet, con botón interactivo de reintento.",
                    alignment: AlignmentType.CENTER,
                    spacing: { after: 250 },
                    run: { italic: true, size: 18, color: mutedText }
                }),

                // SUBSECCIÓN 3.4: RECUPERACIÓN DE RED
                new Paragraph({
                    text: "3.4. Reconexión y Recuperación Exitosa",
                    heading: HeadingLevel.HEADING_3,
                    spacing: { before: 150, after: 100 },
                    run: { color: secondaryColor, bold: true, size: 24 }
                }),
                new Paragraph({
                    children: [
                        new TextRun("Tras rehabilitar la conexión de red y pulsar el botón "),
                        new TextRun({ text: "Retry", bold: true }),
                        new TextRun(", se dispara nuevamente el método "),
                        new TextRun({ text: "getMarsPhotos()", bold: true }),
                        new TextRun(", transicionando de inmediato al estado de carga y posteriormente restaurando la galería completa sin requerir el reinicio de la aplicación."),
                    ],
                    spacing: { after: 150 },
                }),
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                        new ImageRun({
                            data: img4,
                            transformation: { width: 260, height: 550 },
                        }),
                    ],
                    spacing: { after: 100 },
                }),
                new Paragraph({
                    text: "Figura 4: Recuperación exitosa del catálogo tras pulsar el botón Retry.",
                    alignment: AlignmentType.CENTER,
                    spacing: { after: 250 },
                    run: { italic: true, size: 18, color: mutedText }
                }),

                // SECCIÓN 4: PRUEBAS UNITARIAS
                new Paragraph({
                    text: "4. Verificación de Pruebas Unitarias Automatizadas",
                    heading: HeadingLevel.HEADING_2,
                    spacing: { before: 200, after: 150 },
                    run: { color: primaryColor, bold: true, size: 28 }
                }),
                new Paragraph({
                    children: [
                        new TextRun("Se actualizaron y ejecutaron satisfactoriamente las pruebas unitarias en el paquete de tests de la aplicación. En particular, "),
                        new TextRun({ text: "MarsViewModelTest.kt", bold: true }),
                        new TextRun(" comprueba que al inicializarse el ViewModel con un repositorio falso ("),
                        new TextRun({ text: "FakeNetworkMarsPhotosRepository", italic: true }),
                        new TextRun("), el estado emitido corresponda idénticamente a "),
                        new TextRun({ text: "MarsUiState.Success(FakeDataSource.photosList)", bold: true }),
                        new TextRun("."),
                    ],
                    spacing: { after: 150 },
                }),
                new Table({
                    width: { size: 100, type: WidthType.PERCENTAGE },
                    rows: [
                        new TableRow({
                            children: [
                                new TableCell({
                                    shading: { type: ShadingType.CLEAR, fill: "1E293B" },
                                    children: [
                                        new Paragraph({
                                            children: [
                                                new TextRun({
                                                    text: "> Task :app:compileDebugUnitTestKotlin UP-TO-DATE\n> Task :app:testDebugUnitTest\n\nBUILD SUCCESSFUL in 32s\n39 actionable tasks: 21 executed, 18 up-to-date",
                                                    color: "38BDF8", // Cyan code highlight
                                                    font: "Consolas",
                                                    size: 18,
                                                })
                                            ]
                                        })
                                    ]
                                })
                            ]
                        })
                    ]
                }),
                new Paragraph({
                    text: "Salida del comando Gradle: Todas las pruebas unitarias pasaron al 100% sin discrepancias.",
                    alignment: AlignmentType.CENTER,
                    spacing: { before: 100, after: 250 },
                    run: { italic: true, size: 18, color: mutedText }
                }),

                // SECCIÓN 5: ENLACE A GITHUB
                new Paragraph({
                    text: "5. Enlace al Repositorio en GitHub",
                    heading: HeadingLevel.HEADING_2,
                    spacing: { before: 200, after: 150 },
                    run: { color: primaryColor, bold: true, size: 28 }
                }),
                new Paragraph({
                    children: [
                        new TextRun("Todo el código fuente implementado, incluyendo la configuración de Gradle, recursos gráficos, pruebas unitarias y capturas de evidencia, se encuentra publicado en el repositorio público de GitHub:"),
                    ],
                    spacing: { after: 150 },
                }),
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                        new TextRun({
                            text: "https://github.com/Pablojg18/mars-photos-coil",
                            bold: true,
                            size: 26,
                            color: secondaryColor,
                            underline: {}
                        }),
                    ],
                    spacing: { after: 250 },
                }),

                // SECCIÓN 6: CONCLUSIONES
                new Paragraph({
                    text: "6. Conclusiones",
                    heading: HeadingLevel.HEADING_2,
                    spacing: { before: 200, after: 150 },
                    run: { color: primaryColor, bold: true, size: 28 }
                }),
                new Paragraph({
                    children: [
                        new TextRun("1. "),
                        new TextRun({ text: "Superioridad de Coil en Compose: ", bold: true }),
                        new TextRun("A diferencia de bibliotecas tradicionales como Glide o Picasso que fueron diseñadas para el sistema de Views basado en XML, Coil está desarrollado íntegramente en Kotlin y aprovecha de forma nativa las corrutinas y el ciclo de recomposición de Jetpack Compose. El uso del composable "),
                        new TextRun({ text: "AsyncImage", bold: true }),
                        new TextRun(" simplifica drásticamente el código necesario para gestionar placeholders, errores y transiciones sin bloquear el hilo principal."),
                    ],
                    spacing: { after: 100 },
                }),
                new Paragraph({
                    children: [
                        new TextRun("2. "),
                        new TextRun({ text: "Rendimiento y Escalabilidad con LazyVerticalGrid: ", bold: true }),
                        new TextRun("La adopción de listas y cuadrículas perezosas (lazy components) es indispensable cuando se manipulan colecciones extensas de imágenes provenientes de la nube. Gracias a la configuración adaptativa de 150 dp por columna, la interfaz se adapta fluidamente a cualquier factor de forma, garantizando 60 fps estables y previniendo fugas de memoria (OutOfMemoryError)."),
                    ],
                    spacing: { after: 100 },
                }),
                new Paragraph({
                    children: [
                        new TextRun("3. "),
                        new TextRun({ text: "Tolerancia a Fallos y Experiencia de Usuario (UX): ", bold: true }),
                        new TextRun("Un aspecto crítico en el desarrollo móvil es el manejo de contingencias de red. Al modelar estados explícitos ("),
                        new TextRun({ text: "Loading, Success y Error", italic: true }),
                        new TextRun(") en una interfaz sellada e incorporar un mecanismo de reintento directo, la aplicación ofrece una experiencia de usuario resiliente, profesional e intuitiva."),
                    ],
                    spacing: { after: 100 },
                }),
                new Paragraph({
                    children: [
                        new TextRun("4. "),
                        new TextRun({ text: "Importancia de la Arquitectura y Fakes para Testing: ", bold: true }),
                        new TextRun("La inyección de dependencias mediante el contenedor de la aplicación y la separación entre la interfaz del repositorio y su implementación concreta permitieron validar el comportamiento del ViewModel de manera instantánea, sin depender de la disponibilidad del servidor real ni de una conexión activa a Internet."),
                    ],
                    spacing: { after: 300 },
                }),
            ],
        },
    ],
});

Packer.toBuffer(doc).then((buffer) => {
    const outputPath = path.resolve(__dirname, "Reporte_Practica_Mars_Photos_Coil.docx");
    fs.writeFileSync(outputPath, buffer);
    console.log("Documento DOCX generado exitosamente en:", outputPath);
});
