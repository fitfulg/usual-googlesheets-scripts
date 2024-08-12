const cellStyles = {
    "A1": {
        value: {
            "English": "BEHAVIOR PATTERNS",
            "Spanish": "PATRONES DE CONDUCTA",
            "Catalan": "PATRONS DE CONDUCTA"
        },
        fontWeight: "bold",
        fontColor: "#FFFFFF",
        backgroundColor: "#000000",
        alignment: "center"
    },
    "B1": {
        value: {
            "English": "TOMORROW",
            "Spanish": "MAÑANA",
            "Catalan": "DEMÀ"
        },
        fontWeight: "bold",
        fontColor: "#FFFFFF",
        backgroundColor: "#b5a642",
        alignment: "center"
    },
    "B3": {
        value: {
            "English": "WEEK",
            "Spanish": "SEMANA",
            "Catalan": "SETMANA"
        },
        fontWeight: "bold",
        fontColor: "#FFFFFF",
        backgroundColor: "#b5a642",
        alignment: "center"
    },
    "B8": {
        value: {
            "English": "MONTH",
            "Spanish": "MES",
            "Catalan": "MES"
        },
        fontWeight: "bold",
        fontColor: "#FFFFFF",
        backgroundColor: "#b5a642",
        alignment: "center"
    },
    "F1": {
        value: {
            "English": "IDEAS AND PLANS",
            "Spanish": "IDEAS Y PLANES",
            "Catalan": "IDEES I PLANS"
        },
        fontWeight: "bold",
        fontColor: "#000000",
        backgroundColor: "#FFC0CB",
        alignment: "center"
    },
    "G1": {
        value: {
            "English": "EYES ON",
            "Spanish": "ATENTO A",
            "Catalan": "ATENT A"
        },
        fontWeight: "bold",
        fontColor: "#000000",
        backgroundColor: "#b7b7b7",
        alignment: "center"
    },
    "H1": {
        value: {
            "English": "IN QUARANTINE",
            "Spanish": "EN CUARENTENA",
            "Catalan": "EN QUARANTENA"
        },
        fontWeight: "bold",
        fontColor: "#FF0000",
        backgroundColor: null,
        alignment: "center"
    },
    "C1": {
        value: {
            "English": "HIGH PRIORITY",
            "Spanish": "PRIORIDAD ALTA",
            "Catalan": "PRIORITAT ALTA"
        },
        limit: {
            "English": 10,
            "Spanish": 10,
            "Catalan": 10
        },
        priority: {
            "English": "HIGH PRIORITY",
            "Spanish": "PRIORIDAD ALTA",
            "Catalan": "PRIORITAT ALTA"
        },
        fontWeight: "bold",
        fontColor: null,
        backgroundColor: "#fce5cd",
        alignment: "center"
    },
    "D1": {
        value: {
            "English": "MEDIUM PRIORITY",
            "Spanish": "PRIORIDAD MEDIA",
            "Catalan": "PRIORITAT MITJANA"
        },
        limit: {
            "English": 20,
            "Spanish": 20,
            "Catalan": 20
        },
        priority: {
            "English": "MEDIUM PRIORITY",
            "Spanish": "PRIORIDAD MEDIA",
            "Catalan": "PRIORITAT MITJANA"
        },
        fontWeight: "bold",
        fontColor: null,
        backgroundColor: "#fff2cc",
        alignment: "center"
    },
    "E1": {
        value: {
            "English": "LOW PRIORITY",
            "Spanish": "PRIORIDAD BAJA",
            "Catalan": "PRIORITAT BAIXA"
        },
        limit: {
            "English": 20,
            "Spanish": 20,
            "Catalan": 20
        },
        priority: {
            "English": "LOW PRIORITY",
            "Spanish": "PRIORIDAD BAJA",
            "Catalan": "PRIORITAT BAIXA"
        },
        fontWeight: "bold",
        fontColor: null,
        backgroundColor: "#d9ead3",
        alignment: "center"
    }
};


const exampleTexts = {
    "A": {
        text: {
            "English": "Example: Do it with fear but do it.",
            "Spanish": "Ejemplo: Hazlo con miedo pero hazlo.",
            "Catalan": "Exemple: Fes-ho si cal amb por però fes-ho."
        },
        color: "#FFFFFF"
    },
    "B": {
        text: {
            "English": "Example: 45min of cardio",
            "Spanish": "Ejemplo: 45min de cardio",
            "Catalan": "Exemple: 45min de cardio"
        },
        color: "#A9A9A9"
    },
    "C": {
        text: {
            "English": "Example: Join that gym club",
            "Spanish": "Ejemplo: Apuntate al gym",
            "Catalan": "Exemple: Apunta't al gym"
        },
        color: "#A9A9A9"
    },
    "D": {
        text: {
            "English": "Example: Submit that pending data science task.",
            "Spanish": "Ejemplo: Entrega esa tarea pendiente de ciencia de datos.",
            "Catalan": "Exemple: Lliura aquella tasca pendent de ciència de dades."
        },
        color: "#A9A9A9"
    },
    "E": {
        text: {
            "English": "Example: Buy a new mattress.",
            "Spanish": "Ejemplo: Compra un nuevo colchón.",
            "Catalan": "Exemple: Compra un nou matalàs."
        },
        color: "#A9A9A9"
    },
    "F": {
        text: {
            "English": "Example: Santiago route.",
            "Spanish": "Ejemplo: Ruta de Santiago.",
            "Catalan": "Exemple: Ruta de Santiago."
        },
        color: "#A9A9A9"
    },
    "G": {
        text: {
            "English": "Example: Change front brake pad at 44500km",
            "Spanish": "Ejemplo: Cambia la pastilla de freno delantera a los 44500km",
            "Catalan": "Exemple: Canvia la pastilla de fren davanter als 44500km"
        },
        color: "#FFFFFF"
    },
    "H": {
        text: {
            "English": "Example: Join that Crossfit club",
            "Spanish": "Ejemplo: Únete al club de Crossfit",
            "Catalan": "Exemple: Uneix-te al club de Crossfit"
        },
        color: "#A9A9A9"
    }
};


const dateColorConfig = {
    C: { warning: 7, danger: 30, warningColor: '#FFA500', dangerColor: '#FF0000', defaultColor: '#A9A9A9' }, // 1 week, 1 month
    D: { warning: 15, danger: 180, warningColor: '#FFA500', dangerColor: '#FF0000', defaultColor: '#A9A9A9' },
    E: { warning: 180, danger: 365, warningColor: '#FFA500', dangerColor: '#FF0000', defaultColor: '#A9A9A9' },
    F: { warning: 180, danger: 365, warningColor: '#FFA500', dangerColor: '#FF0000', defaultColor: '#A9A9A9' },
    G: { warning: 0, danger: 0, warningColor: '#A9A9A9', dangerColor: '#A9A9A9', defaultColor: '#A9A9A9' }, // Always default
    H: { warning: 0, danger: 0, warningColor: '#FF0000', dangerColor: '#FF0000', defaultColor: '#FF0000' } // Always red
};

const languages = {
    English: 'English',
    Spanish: 'Spanish',
    Catalan: 'Catalan'
};

const menuLanguage = [
    {
        title: {
            English: 'Language',
            Spanish: 'Idioma',
            Catalan: 'Idioma'
        },
        items: {
            setLanguageEnglish: {
                English: 'English',
                Spanish: 'Inglés',
                Catalan: 'Anglès'
            },
            setLanguageSpanish: {
                English: 'Spanish',
                Spanish: 'Español',
                Catalan: 'Espanyol'
            },
            setLanguageCatalan: {
                English: 'Catalan',
                Spanish: 'Catalán',
                Catalan: 'Català'
            }
        }
    }
]
const menuTodoSheet = [
    {
        title: {
            English: 'TODO sheet',
            Spanish: 'Hoja TODO',
            Catalan: 'Full de TODO'
        },
        items: {
            restoreDefaultTodoTemplate: {
                English: 'RESTORE DEFAULT TODO TEMPLATE',
                Spanish: 'RESTAURAR PLANTILLA POR DEFECTO',
                Catalan: 'RESTAURAR PLANTILLA PER DEFECTE'
            },
            restoreCellBackgroundColors: {
                English: 'RESTORE Cell Background Colors',
                Spanish: 'RESTAURAR Colores de Fondo de Celda',
                Catalan: 'RESTAURAR Colors de Fons de Cel·la'
            },
            addCheckboxesToSelectedCells: {
                English: 'Add Checkboxes to Selected Cells',
                Spanish: 'Añadir Casillas a las Celdas Seleccionadas',
                Catalan: 'Afegir Caselles a les Cel·les Seleccionades'
            },
            markCheckboxInSelectedCells: {
                English: 'Mark Checkbox in Selected Cells',
                Spanish: 'Marcar Casilla en las Celdas Seleccionadas',
                Catalan: 'Marcar Casella a les Cel·les Seleccionades'
            },
            markAllCheckboxesInSelectedCells: {
                English: 'Mark All Checkboxes in Selected Cells',
                Spanish: 'Marcar Todas las Casillas en las Celdas Seleccionadas',
                Catalan: 'Marcar Totes les Caselles a les Cel·les Seleccionades'
            },
            restoreCheckboxes: {
                English: 'Restore Checkboxes',
                Spanish: 'Restaurar Casillas',
                Catalan: 'Restaurar Caselles'
            },
            removeAllCheckboxesInSelectedCells: {
                English: 'Remove All Checkboxes in Selected Cells',
                Spanish: 'Eliminar Todas las Casillas en las Celdas Seleccionadas',
                Catalan: 'Eliminar Totes les Caselles a les Cel·les Seleccionades'
            },
            saveSnapshot: {
                English: 'Save Snapshot',
                Spanish: 'Guardar Instantánea',
                Catalan: 'Guardar Instantània'
            },
            restoreSnapshot: {
                English: 'Restore Snapshot',
                Spanish: 'Restaurar Instantánea',
                Catalan: 'Restaurar Instantània'
            },
            createPieChart: {
                English: 'Create Pie Chart',
                Spanish: 'Crear Gráfico Circular',
                Catalan: 'Crear Gràfic Circular'
            },
            deletePieCharts: {
                English: 'Delete Pie Charts',
                Spanish: 'Eliminar Gráficos Circulares',
                Catalan: 'Eliminar Gràfics Circulars'
            },
            versionAndFeatureDetails: {
                English: 'Version and feature details',
                Spanish: 'Detalles de Versión y Funcionalidades',
                Catalan: 'Detalls de Versió i Funcionalitats'
            },
            logHelloWorld: {
                English: 'Log Hello World',
                Spanish: 'Registrar Hola Mundo',
                Catalan: 'Registrar Hola Món'
            }
        }
    }]

const menuCustomFormats = [
    {
        title: {
            English: 'Custom Formats',
            Spanish: 'Formatos Personalizados',
            Catalan: 'Formats Personalitzats'
        },
        items: {
            applyFormat: {
                English: 'Apply Format',
                Spanish: 'Aplicar Formato',
                Catalan: 'Aplicar Format'
            },
            applyFormatToAll: {
                English: 'Apply Format to All',
                Spanish: 'Aplicar Formato a Todo',
                Catalan: 'Aplicar Format a Tot'
            }
        }
    }]

const menus = [
    {
        config: menuTodoSheet,
        items: [
            { key: 'restoreDefaultTodoTemplate', separatorAfter: false },
            { key: 'restoreCellBackgroundColors', separatorAfter: true },
            { key: 'addCheckboxesToSelectedCells', separatorAfter: false },
            { key: 'markCheckboxInSelectedCells', separatorAfter: false },
            { key: 'markAllCheckboxesInSelectedCells', separatorAfter: false },
            { key: 'restoreCheckboxes', separatorAfter: false },
            { key: 'removeAllCheckboxesInSelectedCells', separatorAfter: true },
            { key: 'saveSnapshot', separatorAfter: false },
            { key: 'restoreSnapshot', separatorAfter: true },
            { key: 'createPieChart', separatorAfter: false },
            { key: 'deletePieCharts', separatorAfter: true },
            { key: 'versionAndFeatureDetails', separatorAfter: false },
            { key: 'logHelloWorld', separatorAfter: false }
        ],
        suffix: ''
    },
    {
        config: menuCustomFormats,
        items: [
            { key: 'applyFormat', separatorAfter: false },
            { key: 'applyFormatToAll', separatorAfter: false }
        ],
        suffix: ''
    },
    {
        config: menuLanguage,
        items: [
            { key: 'setLanguageEnglish', separatorAfter: false },
            { key: 'setLanguageSpanish', separatorAfter: false },
            { key: 'setLanguageCatalan', separatorAfter: false }
        ],
        suffix: ''
    }
];

const toastMessages = {
    loading: {
        English: 'Data is loading...\n Please wait.',
        Spanish: 'Cargando datos...\n Por favor espera.',
        Catalan: "S'estan carregant les dades...\n Si us plau, espera."
    },
    updateComplete: {
        English: 'Update Complete!',
        Spanish: 'Actualización completada!',
        Catalan: 'Actualització completada!'
    }
};


if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        cellStyles,
        exampleTexts,
        dateColorConfig,
        languages,
        menuLanguage,
        menuTodoSheet,
        menuCustomFormats,
        menus,
        toastMessages
    }
}