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
        fontWeight: "bold",
        fontColor: null,
        backgroundColor: "#fff2cc",
        alignment: "center"
    },
    "E1": {
        value: {
            "English": "LOW PRIORITY",
            "Spanish": "BAJA PRIORIDAD",
            "Catalan": "BAIXA PRIORITAT"
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
    D: { warning: 90, danger: 180, warningColor: '#FFA500', dangerColor: '#FF0000', defaultColor: '#A9A9A9' },
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

if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        cellStyles,
        exampleTexts,
        dateColorConfig,
        languages
    }
}