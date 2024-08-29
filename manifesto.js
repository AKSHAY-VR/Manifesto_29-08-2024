let candidates = [];
let topics = [];
let topicColors = {};
let subtopicColors = {};

function loadExcelFromGoogleSheet(sheetUrl) {
    fetch(sheetUrl)
        .then(response => response.arrayBuffer())
        .then(data => {
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

            processExcelData(jsonData);
        })
        .catch(error => console.error('Error loading Excel file from Google Sheets:', error));
}

function processExcelData(data) {
    candidates = [];
    topics = new Set();

    data.forEach(row => {
        const candidateName = row.name;
        const image = row.image;
        const partySymbol = row.partySymbol;
        const title = row.title;
        const description = row.description;
        const page = row.page;
        const link = row.link;
        const topic = row.topic;
        const subtopic = row.subtopic || '';

        if (title !== "-" && title.trim() !== "") {
            topics.add(topic);
        }

        let candidate = candidates.find(c => c.name === candidateName);
        if (!candidate) {
            candidate = {
                name: candidateName,
                image: image,
                partySymbol: partySymbol,
                promises: {},
                titleCount: 0,
                linkCount: 0
            };
            candidates.push(candidate);
        }

        if (title !== "-" && title.trim() !== "") {
            candidate.titleCount++;
        }

        if (link && link.trim() !== '') {
            candidate.linkCount++;
        }

        if (!candidate.promises[topic]) {
            candidate.promises[topic] = {};
        }

        if (!candidate.promises[topic][subtopic]) {
            candidate.promises[topic][subtopic] = [];
        }

        candidate.promises[topic][subtopic].push({ title, description, page, link });
    });

    candidates.forEach(candidate => {
        Object.keys(candidate.promises).forEach(topic => {
            const sortedSubtopics = Object.keys(candidate.promises[topic]).sort((a, b) => parseInt(a) - parseInt(b));
            const sortedPromises = {};

            sortedSubtopics.forEach(subtopic => {
                sortedPromises[subtopic] = candidate.promises[topic][subtopic];
            });

            candidate.promises[topic] = sortedPromises;
        });
    });

    topics = Array.from(topics);

    assignColorsToTopics(topics);
    assignColorsToSubtopics();

    populateCandidateCardsSection(candidates);
    createSectionsForTopics(topics);
    populateCandidateCheckboxes(candidates);
    populateTopicCheckboxes(topics);
    updateDisplay();
}

function populateCandidateCardsSection(candidates) {
    const candidateCardsContainer = d3.select("#candidate-cards");
    candidateCardsContainer.html('');

    candidates.forEach(candidate => {
        const candidateCard = candidateCardsContainer.append("div")
            .attr("class", "new-candidate-card");

        candidateCard.append("img")
            .attr("src", candidate.image)
            .attr("alt", candidate.name);

        candidateCard.append("h3")
            .text(candidate.name);

        candidateCard.append("p")
            .attr("class", "summary-stats")
            .text(`Rationales per Promise: ${candidate.linkCount}/${candidate.titleCount}`);

        candidateCard.append("div")
            .attr("class", "party-symbol")
            .append("img")
            .attr("src", candidate.partySymbol)
            .attr("alt", `${candidate.name} party symbol`);
    });

    equalizeCandidateCardWidths();
}

function equalizeCandidateCardWidths() {
    const cards = document.querySelectorAll('.new-candidate-card');
    let maxWidth = 0;

    cards.forEach(card => {
        const width = card.getBoundingClientRect().width;
        if (width > maxWidth) {
            maxWidth = width;
        }
    });

    cards.forEach(card => {
        card.style.width = `${maxWidth}px`;
    });
}

function createSectionsForTopics(topics) {
    const contentContainer = d3.select("#content");

    topics.forEach(topic => {
        const topicSection = contentContainer.append("div")
            .attr("id", `${topic.toLowerCase().replace(/\s+/g, '-')}-section`)
            .attr("class", "topic-section");

        topicSection.append("div")
            .attr("class", "section-title-container")
            .append("h2")
            .attr("class", "section-title")
            .text(topic);

        topicSection.append("div")
            .attr("class", "candidate-card-container")
            .attr("id", `${topic.toLowerCase().replace(/\s+/g, '-')}`);
    });
}

function populateCandidateCheckboxes(candidates) {
    const candidateCheckboxesContainer = d3.select("#candidate-checkboxes");
    candidateCheckboxesContainer.html('');

    const allLabel = candidateCheckboxesContainer.append("label");
    allLabel.append("input")
        .attr("type", "checkbox")
        .attr("value", "all")
        .on("change", function() {
            const isChecked = this.checked;
            candidateCheckboxesContainer.selectAll("input").property("checked", isChecked);
            updateDisplay();
        });
    allLabel.append("span").text("All");
    allLabel.append("br");

    candidates.forEach(candidate => {
        const checkboxLabel = candidateCheckboxesContainer.append("label");
        checkboxLabel.append("input")
            .attr("type", "checkbox")
            .attr("value", candidate.name)
            .on("change", updateDisplay);
        checkboxLabel.append("span").text(candidate.name);
        checkboxLabel.append("br");
    });
}

function populateTopicCheckboxes(topics) {
    const topicCheckboxesContainer = d3.select("#topic-checkboxes");
    topicCheckboxesContainer.html('');

    const allLabel = topicCheckboxesContainer.append("label");
    allLabel.append("input")
        .attr("type", "checkbox")
        .attr("value", "all")
        .on("change", function() {
            const isChecked = this.checked;
            topicCheckboxesContainer.selectAll("input").property("checked", isChecked);
            updateDisplay();
        });
    allLabel.append("span").text("All");
    allLabel.append("br");

    topics.forEach(topic => {
        const checkboxLabel = topicCheckboxesContainer.append("label");
        checkboxLabel.append("input")
            .attr("type", "checkbox")
            .attr("value", topic)
            .on("change", updateDisplay);
        checkboxLabel.append("span").text(topic);
        checkboxLabel.append("br");
    });
}

function updateDisplay() {
    const selectedCandidates = [];
    d3.selectAll("#candidate-checkboxes input:checked").each(function() {
        if (this.value !== "all") {
            selectedCandidates.push(this.value);
        }
    });

    const selectedTopics = [];
    d3.selectAll("#topic-checkboxes input:checked").each(function() {
        if (this.value !== "all") {
            selectedTopics.push(this.value);
        }
    });

    let filteredCandidates = candidates;

    if (selectedCandidates.length > 0) {
        filteredCandidates = filteredCandidates.filter(candidate => selectedCandidates.includes(candidate.name));
    }

    topics.forEach(topic => {
        const sectionId = `#${topic.toLowerCase().replace(/\s+/g, '-')}`;
        d3.select(sectionId).selectAll(".candidate-card").remove();
    });

    filteredCandidates.forEach(candidate => {
        Object.keys(candidate.promises).forEach(topic => {
            if (selectedTopics.length === 0 || selectedTopics.includes(topic)) {
                createOrUpdateCandidateCard(`#${topic.toLowerCase().replace(/\s+/g, '-')}`, candidate, candidate.promises[topic], topic.toLowerCase().replace(/\s+/g, '-'));
            }
        });
    });

    topics.forEach(topic => {
        const sectionId = `#${topic.toLowerCase().replace(/\s+/g, '-')}-section`;
        if (d3.select(`#${topic.toLowerCase().replace(/\s+/g, '-')}`).selectAll(".candidate-card").empty()) {
            d3.select(sectionId).classed("hidden", true);
        } else {
            d3.select(sectionId).classed("hidden", false);
        }
    });

    equalizePromiseHeights();
}

function createOrUpdateCandidateCard(containerId, candidate, promisesByTopic, topicClass) {
    const container = d3.select(containerId);

    let candidateCard = container.select(`.candidate-card[data-candidate="${candidate.name}"]`);

    if (candidateCard.empty()) {
        candidateCard = container.append("div")
            .attr("class", `candidate-card ${topicClass}`)
            .attr("data-candidate", candidate.name);

        candidateCard.append("img")
            .attr("src", candidate.image)
            .attr("alt", candidate.name);

        candidateCard.append("h3")
            .text(candidate.name);

        candidateCard.append("p")
            .attr("class", "summary-stats")
            .text(`Rationales per Promise: 0/0`);

        candidateCard.append("div")
            .attr("class", "party-symbol")
            .append("img")
            .attr("src", candidate.partySymbol)
            .attr("alt", `${candidate.name} party symbol`);

        candidateCard.append("div")
            .attr("class", "promises-container");
    }

    let topicTitleCount = 0;
    let topicLinkCount = 0;

    Object.keys(promisesByTopic).forEach(subtopic => {
        const subtopicPromises = promisesByTopic[subtopic];

        subtopicPromises.forEach(promise => {
            if (promise.title && promise.title.trim() !== "" && promise.title !== "-") {
                topicTitleCount++;
            }
            if (promise.link && promise.link.trim() !== "") {
                topicLinkCount++;
            }
        });
    });

    candidateCard.select(".summary-stats")
        .text(`Rationales per Promise: ${topicLinkCount}/${topicTitleCount}`);

    Object.keys(promisesByTopic).forEach(subtopic => {
        const subtopicPromises = promisesByTopic[subtopic];

        subtopicPromises.forEach(promise => {
            const promiseContainer = candidateCard.select(".promises-container")
                .append("div")
                .attr("class", `promise-container ${topicClass}`)
                .style("background-color", getColorBySubtopic(subtopic))
                .style("margin-bottom", "10px")
                .style("padding", "10px")
                .style("border-radius", "8px");

            promiseContainer.append("div")
                .attr("class", `promise ${topicClass}`)
                .html(`<p class="promise-title">${promise.title}</p><p class="promise-page">Page: ${promise.page}</p><p>${promise.description}</p>`);

            if (promise.link && promise.link.trim() !== '') {
                promiseContainer.append("button")
                    .attr("class", "reference-button")
                    .text("Rationale")
                    .on("click", () => {
                        window.open(promise.link, '_blank');
                    });
            }
        });
    });
}

function equalizePromiseHeights() {
    topics.forEach(topic => {
        const promiseContainers = d3.select(`#${topic.toLowerCase().replace(/\s+/g, '-')}`)
                                    .selectAll('.promise-container')
                                    .nodes();

        let maxHeight = 0;

        promiseContainers.forEach(container => {
            const height = container.getBoundingClientRect().height;
            if (height > maxHeight) {
                maxHeight = height;
            }
        });

        promiseContainers.forEach(container => {
            container.style.height = `${maxHeight}px`;
        });
    });
}

function assignColorsToTopics(topics) {
    topics.forEach(topic => {
        if (!topicColors[topic]) {
            topicColors[topic] = generateRandomColor();
        }
    });
}

function assignColorsToSubtopics() {
    candidates.forEach(candidate => {
        Object.keys(candidate.promises).forEach(topic => {
            Object.keys(candidate.promises[topic]).forEach(subtopic => {
                if (!subtopicColors[subtopic] && subtopic !== '') {
                    subtopicColors[subtopic] = generateRandomColor();
                }
            });
        });
    });
}

function getColorBySubtopic(subtopic) {
    return subtopicColors[subtopic] || "#ffffff";
}

let hueShift = 0;

function generateRandomColor() {
    const saturation = 60;
    const lightness = 80;

    hueShift += 137.5;
    const hue = hueShift % 360;

    return `hsl(${hue}, ${saturation}%, ${lightness}%)`;
}

document.addEventListener("click", function(event) {
    const filterContainer = document.querySelector(".filter-container");
    if (!filterContainer.contains(event.target)) {
        closeAllCheckboxes();
    }
});

function closeAllCheckboxes() {
    document.querySelectorAll(".checkboxes").forEach(checkboxContainer => {
        checkboxContainer.style.display = "none";
    });
}

function toggleCheckboxes(id) {
    const checkboxes = document.getElementById(id);
    if (checkboxes.style.display === "block") {
        checkboxes.style.display = "none";
    } else {
        closeAllCheckboxes();
        checkboxes.style.display = "block";
    }
}

loadExcelFromGoogleSheet('https://docs.google.com/spreadsheets/d/1_q6IaRErhJnPM_pTwiI7pGvOTJ4PJJRMTaz4zr-rmio/export?format=xlsx');
