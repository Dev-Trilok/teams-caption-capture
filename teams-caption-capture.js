/**
 * teams-caption-capture.js
 *
 * 🧠 Description: A JavaScript script to capture live closed captions from Microsoft Teams (Web).
 * 💡 How it works: It observes the Teams DOM and stores unique speaker + caption pairs in localStorage.
 * 📥 To export as .yaml, run: downloadYAML();
 *
 * 🔗 GitHub: https://github.com/Dev-Trilok/teams-caption-capture
 * 📅 Created: 2025-07-15
 * 👤 Author: Triloknath Nalawade
 * 🌐 License: MIT
 */


localStorage.removeItem("transcripts");
const transcriptArray = [];
const seen = new Set();
let transcriptIdCounter = 0;

function checkTranscripts() {
    const transcripts = document.querySelectorAll(".fui-ChatMessageCompact");

    transcripts.forEach(transcript => {
        const authorElement = transcript.querySelector('[data-tid="author"]');
        const textElement = transcript.querySelector('[data-tid="closed-caption-text"]');

        if (!authorElement || !textElement) return;

        const Name = authorElement.innerText.trim();
        const Text = textElement.innerText.trim();
        const Time = new Date().toISOString().replace('T', ' ').slice(0, -5);
        const key = `${Name}||${Text}`;

        if (!seen.has(key) && Text.length > 0) {
            seen.add(key);
            const ID = `caption_${transcriptIdCounter++}`;
            console.log({ Name, Text, Time, ID });
            transcriptArray.push({ Name, Text, Time, ID });
        }
    });

    localStorage.setItem('transcripts', JSON.stringify(transcriptArray));
}

function initObserver() {
    const target = document.querySelector('.fui-ChatMessageCompact')?.parentNode?.parentNode;

    if (!target) {
        console.log("⏳ Waiting for captions container...");
        setTimeout(initObserver, 2000);
        return;
    }

    const observer = new MutationObserver(checkTranscripts);
    observer.observe(target, {
        childList: true,
        subtree: true,
        characterData: true
    });

    console.log("✅ Caption observer started.");
    checkTranscripts();
    setInterval(checkTranscripts, 10000);
}

function downloadYAML() {
    let transcripts = JSON.parse(localStorage.getItem('transcripts')) || [];

    if (!transcripts.length) {
        alert("No captions found.");
        return;
    }

    transcripts = transcripts.map(({ ID, ...rest }) => rest);

    let yamlTranscripts = '';
    transcripts.forEach(t => {
        yamlTranscripts += `Name: ${t.Name}\nText: ${t.Text}\n----\n`;
    });

    let title = document.title.replace("__Microsoft_Teams", '').replace(/[^a-z0-9 ]/gi, '');
    const fileName = "transcript - " + title.trim() + ".yaml";

    const dataStr = "data:text/yaml;charset=utf-8," + encodeURIComponent(yamlTranscripts);
    const downloadAnchorNode = document.createElement('a');
    downloadAnchorNode.setAttribute("href", dataStr);
    downloadAnchorNode.setAttribute("download", fileName);
    document.body.appendChild(downloadAnchorNode);
    downloadAnchorNode.click();
    downloadAnchorNode.remove();
}

console.log("📽️ Teams Caption Capture is running!");
console.log("📥 To download, run: downloadYAML()");
initObserver();
