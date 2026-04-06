const form = document.getElementById("projectForm");
const output = document.getElementById("output");
const downloadBtn = document.getElementById("downloadBtn");
const wordBtn = document.getElementById("wordBtn");

let latestJSON = null;

form.addEventListener("submit", async (e) => {
    e.preventDefault();

    output.textContent = "Generating...";

    const projectName = document.getElementById("projectName").value;
    const industry = document.getElementById("industry").value;
    const duration = document.getElementById("duration").value;
    const teamSize = document.getElementById("teamSize").value;
    const objective = document.getElementById("objective").value;
    const features = document.getElementById("features").value;

    const prompt = `
Act as a Senior Project Manager.

Return STRICT JSON only in this structure:

{
  "project_charter": {
    "objective": "",
    "scope_in": [],
    "scope_out": [],
    "deliverables": [],
    "success_criteria": []
  },
  "sprint_plan": [],
  "risk_register": []
}

Project Name: ${projectName}
Industry: ${industry}
Duration: ${duration}
Team Size: ${teamSize}
Objective: ${objective}
Key Features: ${features}
`;

    try {
        const response = await fetch("https://api.openai.com/v1/chat/completions", {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                "Authorization": "Bearer sk-proj-mNFOCkv7uceSXo2lIiknhCFV9rSsKcGoU6Vql8xIBjadKmnmWEtJLZq0HfvPxoln9qviO2fSRET3"
            },
            body: JSON.stringify({
                model: "gpt-4o-mini",
                messages: [{ role: "user", content: prompt }]
            })
        });

        const data = await response.json();

        if (!response.ok) {
            output.textContent = "API Error:\n" + JSON.stringify(data, null, 2);
            return;
        }

        const cleaned = data.choices[0].message.content
            .replace(/```json|```/g, "")
            .trim();

        const jsonData = JSON.parse(cleaned);
        latestJSON = jsonData;

        // Save to localStorage
        const existing = JSON.parse(localStorage.getItem("projects")) || [];
        existing.push(jsonData);
        localStorage.setItem("projects", JSON.stringify(existing));

        output.textContent = JSON.stringify(jsonData, null, 2);

    } catch (error) {
        output.textContent = "Error:\n" + error.message;
    }
});


/* ---------- Download JSON ---------- */
downloadBtn.addEventListener("click", () => {
    if (!latestJSON) return alert("Generate first!");

    const blob = new Blob(
        [JSON.stringify(latestJSON, null, 2)],
        { type: "application/json" }
    );

    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "project_documentation.json";
    a.click();
});


/* ---------- Download Word ---------- */
wordBtn.addEventListener("click", () => {

    if (!latestJSON) return alert("Generate first!");

    const { Document, Packer, Paragraph, HeadingLevel } = docx;

    const children = [];

    // Project Charter
    children.push(
        new Paragraph({
            text: "Project Charter",
            heading: HeadingLevel.HEADING_1
        })
    );

    children.push(
        new Paragraph("Objective: " + latestJSON.project_charter.objective)
    );

    // Scope In
    children.push(new Paragraph({ text: "Scope In:", heading: HeadingLevel.HEADING_2 }));
    latestJSON.project_charter.scope_in.forEach(item => {
        children.push(new Paragraph("- " + item));
    });

    // Deliverables
    children.push(new Paragraph({ text: "Deliverables:", heading: HeadingLevel.HEADING_2 }));
    latestJSON.project_charter.deliverables.forEach(item => {
        children.push(new Paragraph("- " + item));
    });

    // Sprint Plan
    children.push(
        new Paragraph({
            text: "Sprint Plan",
            heading: HeadingLevel.HEADING_1
        })
    );

    latestJSON.sprint_plan.forEach((sprint, index) => {
        children.push(new Paragraph(`Sprint ${index + 1}: ${sprint.sprint_name}`));
    });

    // Risk Register
    children.push(
        new Paragraph({
            text: "Risk Register",
            heading: HeadingLevel.HEADING_1
        })
    );

    latestJSON.risk_register.forEach(risk => {
        children.push(
            new Paragraph(
                `Risk: ${risk.risk} | Impact: ${risk.impact} | Mitigation: ${risk.mitigation}`
            )
        );
    });

    const doc = new Document({
        sections: [{ children }]
    });

    Packer.toBlob(doc).then(blob => {
        saveAs(blob, "Project_Documentation.docx");
    });
});