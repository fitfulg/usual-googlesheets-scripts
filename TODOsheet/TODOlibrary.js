const cellStyles = {
    "A1": { value: "QUICK PATTERNS", fontWeight: "bold", fontColor: "#FFFFFF", backgroundColor: "#000000", alignment: "center" },
    "B1": { value: "TOMORROW", fontWeight: "bold", fontColor: "#FFFFFF", backgroundColor: "#b5a642", alignment: "center" },
    "B3": { value: "WEEK", fontWeight: "bold", fontColor: "#FFFFFF", backgroundColor: "#b5a642", alignment: "center" },
    "B8": { value: "MONTH", fontWeight: "bold", fontColor: "#FFFFFF", backgroundColor: "#b5a642", alignment: "center" },
    "F1": { value: "💡IDEAS AND PLANS", fontWeight: "bold", fontColor: "#000000", backgroundColor: "#FFC0CB", alignment: "center" },
    "G1": { value: "👀 EYES ON", fontWeight: "bold", fontColor: "#000000", backgroundColor: "#b7b7b7", alignment: "center" },
    "H1": { value: "IN QUARANTINE BEFORE BEING CANCELED", fontWeight: "bold", fontColor: "#FF0000", backgroundColor: null, alignment: "center" },
    "C1": { value: "HIGH PRIORITY", fontWeight: "bold", fontColor: null, backgroundColor: "#fce5cd", alignment: "center" },
    "D1": { value: "MEDIUM PRIORITY", fontWeight: "bold", fontColor: null, backgroundColor: "#fff2cc", alignment: "center" },
    "E1": { value: "LOW PRIORITY", fontWeight: "bold", fontColor: null, backgroundColor: "#d9ead3", alignment: "center" }
};

const exampleTexts = {
    "A": { text: "Example: Do it with fear but do it.", color: "#FFFFFF" },
    "B": { text: "Example: 45min of cardio", color: "#A9A9A9" },
    "C": { text: "Example: Join that gym club", color: "#A9A9A9" },
    "D": { text: "Example: Submit that pending data science task.", color: "#A9A9A9" },
    "E": { text: "Example: Buy a new mattress.", color: "#A9A9A9" },
    "F": { text: "Example: Santiago route.", color: "#A9A9A9" },
    "G": { text: "Example: Change front brake pad at 44500km", color: "#FFFFFF" },
    "H": { text: "Example: Join that Crossfit club", color: "#A9A9A9" },
};

const dateColorConfig = {
    C: { warning: 7, danger: 30, warningColor: '#FFA500', dangerColor: '#FF0000', defaultColor: '#A9A9A9' }, // 1 week, 1 month
    D: { warning: 90, danger: 180, warningColor: '#FFA500', dangerColor: '#FF0000', defaultColor: '#A9A9A9' }, // 3 months, 6 months
    E: { warning: 180, danger: 365, warningColor: '#FFA500', dangerColor: '#FF0000', defaultColor: '#A9A9A9' }, // 6 months, 1 year
    F: { warning: 180, danger: 365, warningColor: '#FFA500', dangerColor: '#FF0000', defaultColor: '#A9A9A9' }, // 6 months, 1 year
    G: { warning: 0, danger: 0, warningColor: '#A9A9A9', dangerColor: '#A9A9A9', defaultColor: '#A9A9A9' }, // Always default
    H: { warning: 0, danger: 0, warningColor: '#FF0000', dangerColor: '#FF0000', defaultColor: '#FF0000' } // Always red
};