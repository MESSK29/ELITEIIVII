// Handle form submission to add or update marks
document.getElementById("marksForm").addEventListener("submit", async function(event) {
    event.preventDefault();

    const rollNumber = document.getElementById("rollNumber").value;
    const section = document.getElementById("section").value;
    const marks = document.getElementById("marks").value;
    const session = document.querySelector('input[name="session"]:checked');

    if (!session) {
        alert("Please select a session (Morning or Evening).");
        return;
    }

    const sessionType = session.value;
    const studentData = { rollNumber, section };

    if (sessionType === "morning") {
        studentData.morningMarks = marks;
    } else {
        studentData.eveningMarks = marks;
    }

    try {
        const response = await fetch("https://elite27.onrender.com/add-marks", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(studentData)
        });

        if (!response.ok) {
            const errorData = await response.json();
            throw new Error(errorData.message || "Failed to save marks.");
        }

        const result = await response.json();
        alert(result.message);
        document.getElementById("marksForm").reset();
    } catch (error) {
        alert("Error saving marks: " + error.message);
        console.error("Error:", error);
    }
});

// Function to download datastored.xlsx
function downloadResults() {
    const code = document.getElementById("dataCode").value.trim();
    if (code !== "data") {
        alert("Invalid code! Please enter a vaild code.");
        return;
    }
    window.location.href = `https://elite27.onrender.com/download-results?code=${encodeURIComponent(code)}`;
}

// Function to download winners.xlsx
function downloadWinners() {
    const code = document.getElementById("winnersCode").value.trim();
    if (code !== "winner") {
        alert("Invalid code! Please enter enter a vaild code.");
        return;
    }
    window.location.href = `https://elite27.onrender.com/download-winners?code=${encodeURIComponent(code)}`;
}