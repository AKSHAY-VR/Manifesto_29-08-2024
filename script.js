let candidates = [];
let topics = [];
let topicColors = {};  // Object to store colors for each topic
let subtopicColors = {}; // Object to store colors for each subtopic

// Function to load Excel data from a Google Sheets URL
function loadExcelFromGoogleSheet(sheetUrl) {
    fetch(sheetUrl)
        .then(response => response.arrayBuffer())
        .then(data => {
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

            // Process the Excel data
            processExcelData(jsonData);
        })
        .catch(error => console.error('Error loading Excel file from Google Sheets:', error));
}

// Function to process Excel data and update candidates and topics
function processExcelData(data) {
    candidates = [];
    topics = new Set();  // Use a set to handle unique topics automatically

    data.forEach(row => {
        const candidateName = row.name;
        const image = row.image;
        const partySymbol = row.partySymbol;
        const title = row.title;
        const description = row.description;
        const page = row.page;
        const link = row.link;
        const topic = row.topic;
        const subtopic = row.subtopic || ''; // New subtopic column, default to empty string if not provided

        // Only add valid titles (not N/A) to the title count
        if (title !== "-" && title.trim() !== "") {
            topics.add(topic); // Add topic to the set (duplicates are automatically handled)
        }

        let candidate = candidates.find(c => c.name === candidateName);
        if (!candidate) {
            candidate = {
                name: candidateName,
                image: image,
                partySymbol: partySymbol,
                promises: {},
                titleCount: 0, // Initialize title count
                linkCount: 0   // Initialize link count
            };
            candidates.push(candidate);
        }

        // Increment the title count only for valid titles
        if (title !== "-" && title.trim() !== "") {
            candidate.titleCount++;
        }

        // Increment the link count if a valid link exists
        if (link && link.trim() !== '') {
            candidate.linkCount++;
        }

        // Initialize topic if it doesn't exist
        if (!candidate.promises[topic]) {
            candidate.promises[topic] = {};
        }

        // Initialize subtopic if it doesn't exist under the topic
        if (!candidate.promises[topic][subtopic]) {
            candidate.promises[topic][subtopic] = [];
        }

        // Add the promise under the appropriate subtopic
        candidate.promises[topic][subtopic].push({ title, description, page, link });
    });

    // Sort the subtopics numerically (assuming subtopics are numbers)
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

    // Assign colors to topics dynamically
    assignColorsToTopics(topics);
    assignColorsToSubtopics(); // Assign colors to subtopics dynamically

    // Populate the new candidate cards section
    populateCandidateCardsSection(candidates);

    // Create sections for each topic
    createSectionsForTopics(topics);

    // Update the UI with new candidates and topics
    populateCandidateCheckboxes(candidates);
    populateTopicCheckboxes(topics);
    updateDisplay();

    // Create Fantasy 11 page
    createFantasy11Page(candidates);
}

// Function to populate the new candidate cards section
function populateCandidateCardsSection(candidates) {
    const candidateCardsContainer = d3.select("#candidate-cards");
    candidateCardsContainer.html(''); // Clear existing content

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

    // Ensure all candidate cards have the same width
    equalizeCandidateCardWidths();
}

// Function to equalize the width of candidate cards
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

// Function to create sections for each topic
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
    candidateCheckboxesContainer.html(''); // Clear existing content

    // "All" option
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
    topicCheckboxesContainer.html(''); // Clear existing content

    // "All" option
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
        d3.select(sectionId).selectAll(".candidate-card").remove(); // Clear existing cards
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

    // Equalize heights after updating display
    equalizePromiseHeights();
}

// Function to create or update candidate cards with promises grouped by subtopic
function createOrUpdateCandidateCard(containerId, candidate, promisesByTopic, topicClass) {
    const container = d3.select(containerId);

    // Create or select a single card for the candidate
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

        // Add the "Rationals per Promise" summary right after the candidate's name
        candidateCard.append("p")
            .attr("class", "summary-stats")
            .text(`Rationales per Promise: 0/0`);  // Placeholder text, updated below

        candidateCard.append("div")
            .attr("class", "party-symbol")
            .append("img")
            .attr("src", candidate.partySymbol)
            .attr("alt", `${candidate.name} party symbol`);

        // Create the container for promises
        candidateCard.append("div")
            .attr("class", "promises-container");
    }

    // Calculate topic-specific title and link counts
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

    // Update the "Rationals per Promise" summary for this candidate and topic
    candidateCard.select(".summary-stats")
        .text(`Rationales per Promise: ${topicLinkCount}/${topicTitleCount}`);

    // Add or update promises grouped by subtopic
    Object.keys(promisesByTopic).forEach(subtopic => {
        const subtopicPromises = promisesByTopic[subtopic];

        // Add each promise separately, even if they have the same subtopic
        subtopicPromises.forEach(promise => {
            const promiseContainer = candidateCard.select(".promises-container")
                .append("div")
                .attr("class", `promise-container ${topicClass}`)
                .style("background-color", getColorBySubtopic(subtopic)) // Set background color based on the subtopic
                .style("margin-bottom", "10px")
                .style("padding", "10px")
                .style("border-radius", "8px");

            promiseContainer.append("div")
                .attr("class", `promise ${topicClass}`)
                .html(`<p class="promise-title">${promise.title}</p><p class="promise-page">Page: ${promise.page}</p><p>${promise.description}</p>`);

            // Add the "Rationale" button under each promise if there's a valid link
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

// Function to equalize the height of promise containers within each topic
function equalizePromiseHeights() {
    topics.forEach(topic => {
        const promiseContainers = d3.select(`#${topic.toLowerCase().replace(/\s+/g, '-')}`)
                                    .selectAll('.promise-container')
                                    .nodes();

        let maxHeight = 0;

        // Calculate the maximum height of promise containers within this topic
        promiseContainers.forEach(container => {
            const height = container.getBoundingClientRect().height;
            if (height > maxHeight) {
                maxHeight = height;
            }
        });

        // Set the maximum height for all promise containers within this topic
        promiseContainers.forEach(container => {
            container.style.height = `${maxHeight}px`;
        });
    });
}

// Function to assign colors to topics dynamically
function assignColorsToTopics(topics) {
    topics.forEach(topic => {
        if (!topicColors[topic]) {
            topicColors[topic] = generateRandomColor();
        }
    });
}

// Function to assign colors to subtopics dynamically
function assignColorsToSubtopics() {
    // Loop through each candidate and each topic to assign subtopic colors
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

// Function to get color by subtopic
function getColorBySubtopic(subtopic) {
    return subtopicColors[subtopic] || "#ffffff"; // Default to white if the subtopic doesn't have a specific color
}

let hueShift = 0; // Initialize hue shift

// Function to generate a unique lighter color using HSL
function generateRandomColor() {
    const saturation = 60; // Fixed saturation level for vibrancy
    const lightness = 80;  // Fixed lightness level for lighter colors

    hueShift += 137.5; // Golden angle approximation to distribute hues evenly
    const hue = hueShift % 360; // Ensure hue is within 0-360 degrees

    return `hsl(${hue}, ${saturation}%, ${lightness}%)`;
}

// Add event listener to close checkboxes when clicking outside the filter container
document.addEventListener("click", function(event) {
    const filterContainer = document.querySelector(".filter-container");
    if (!filterContainer.contains(event.target)) {
        closeAllCheckboxes();
    }
});

// Function to close all checkboxes
function closeAllCheckboxes() {
    document.querySelectorAll(".checkboxes").forEach(checkboxContainer => {
        checkboxContainer.style.display = "none";
    });
}

// Modify the existing toggleCheckboxes function to close other checkboxes when opening a new one
function toggleCheckboxes(id) {
    const checkboxes = document.getElementById(id);
    if (checkboxes.style.display === "block") {
        checkboxes.style.display = "none";
    } else {
        closeAllCheckboxes();  // Close all other checkboxes
        checkboxes.style.display = "block";
    }
}

// Function to display the Manifesto Promises page
function showManifestoPage() {
    document.getElementById('content').style.display = 'block'; // Show the original content
    document.getElementById('fantasy11-content').style.display = 'none'; // Hide Fantasy 11 content
}

// Function to display the Fantasy 11 page
function showFantasy11Page() {
    document.getElementById('content').style.display = 'none'; // Hide the original content
    document.getElementById('fantasy11-content').style.display = 'block'; // Show Fantasy 11 content
}

// Function to create the Fantasy 11 page content
function createFantasy11Page(candidates) {
    const fantasyContainer = d3.select("#fantasy-candidates-container");
    fantasyContainer.html(''); // Clear any existing content

    candidates.forEach(candidate => {
        const candidateContainer = fantasyContainer.append("div")
            .attr("class", "fantasy-candidate-container");

        candidateContainer.append("h3")
            .attr("class", "fantasy-candidate-title")
            .text(candidate.name);

        // Iterate over each topic for the candidate
        Object.keys(candidate.promises).forEach(topic => {
            const subtopics = candidate.promises[topic];
            Object.keys(subtopics).forEach(subtopic => {
                const promises = subtopics[subtopic];

                // Add promises as selectable options
                promises.forEach(promise => {
                    const promiseContainer = candidateContainer.append("div")
                        .attr("class", "fantasy-promise");

                    promiseContainer.append("input")
                        .attr("type", "checkbox")
                        .attr("value", promise.title);

                    promiseContainer.append("label")
                        .text(`${promise.title} - Page: ${promise.page}`);
                });
            });
        });
    });
}

// Function to display the pop-up with candidate cards based on selected subtopics
function showCandidatePopup() {
    const selectedSubtopics = Array.from(document.getElementById('selected-subtopics').children).map(child => child.innerText);
    const candidateCardsContainer = document.getElementById('candidate-cards');
    candidateCardsContainer.innerHTML = ''; // Clear previous cards

    const candidateMatches = [];

    candidates.forEach(candidate => {
        let matchCount = 0;
        selectedTopics.forEach(topic => {
            selectedSubtopics.forEach(subtopic => {
                if (candidate.promises[topic][subtopic]) {
                    matchCount++;
                }
            });
        });

        if (matchCount > 0) {
            candidateMatches.push({
                name: candidate.name,
                image: candidate.image,
                matchCount: matchCount,
                totalCount: selectedSubtopics.length
            });
        }
    });

    candidateMatches.sort((a, b) => (b.matchCount / b.totalCount) - (a.matchCount / a.totalCount));
    const topCandidates = candidateMatches.slice(0, 3); // Get top 3 candidates

    topCandidates.forEach(candidate => {
        const card = document.createElement('div');
        card.className = 'candidate-card';
        card.innerHTML = `
            <img src="${candidate.image}" alt="${candidate.name}">
            <h3>${candidate.name}</h3>
            <p>Matching Promises: ${candidate.matchCount}/${candidate.totalCount} (${((candidate.matchCount / candidate.totalCount) * 100).toFixed(2)}%)</p>
        `;
        candidateCardsContainer.appendChild(card);
    });

    document.getElementById('result-page').style.display = 'block'; // Show the result page
}

// Function to close the pop-up modal
function closePopup() {
    document.getElementById('candidate-popup').style.display = 'none';
}

// Fantasy 11 game functions
function startFantasyGame() {
    document.getElementById('fantasy11-game-start').style.display = 'none';
    document.getElementById('topic-selection').style.display = 'block';
    populateTopicOptions();
}

function populateTopicOptions() {
    const topicOptionsContainer = document.getElementById('topic-options');
    topicOptionsContainer.innerHTML = ''; // Clear previous options

    topics.forEach(topic => {
        const label = document.createElement('label');
        label.innerHTML = `<input type="checkbox" value="${topic}" onchange="updateTopicSelection()"> ${topic}`;
        topicOptionsContainer.appendChild(label);
    });
}

let selectedTopics = [];

function updateTopicSelection() {
    selectedTopics = Array.from(document.querySelectorAll('#topic-options input:checked')).map(input => input.value);
    document.getElementById('confirm-topics-button').disabled = selectedTopics.length < 3;
}

function confirmTopics() {
    document.getElementById('topic-selection').style.display = 'none';
    document.getElementById('subtopic-selection').style.display = 'block';
    populateSubtopics();
}

function populateSubtopics() {
    const availableSubtopicsContainer = document.getElementById('available-subtopics');
    availableSubtopicsContainer.innerHTML = ''; // Clear previous subtopics

    candidates.forEach(candidate => {
        selectedTopics.forEach(topic => {
            const subtopics = candidate.promises[topic]; // Get subtopics for the topic

            // Display each subtopic's promises
            Object.keys(subtopics).forEach(subtopic => {
                const promises = subtopics[subtopic];
                promises.forEach((promise, index) => {
                    const subtopicBox = document.createElement('div');
                    subtopicBox.className = 'subtopic-box';
                    subtopicBox.id = `promise-${topic}-${subtopic}-${index}`; // Assign a unique id
                    subtopicBox.draggable = true;
                    subtopicBox.ondragstart = drag;
                    subtopicBox.innerHTML = `
                        <p>${promise.description}</p>
                    `;
                    availableSubtopicsContainer.appendChild(subtopicBox);
                });
            });
        });
    });
}


function drop(event) {
    event.preventDefault();
    const id = event.dataTransfer.getData('text');  // Retrieve the id of the dragged element
    const draggedElement = document.getElementById(id);  // Get the exact element by its id

    const targetContainer = event.target.closest('.promise-list-container');

    if (targetContainer && targetContainer !== draggedElement.parentElement) {
        // Move the exact dragged element to the target container
        targetContainer.appendChild(draggedElement);

        // Enable the Finalize button if dropping into the bottom container
        if (targetContainer.id === 'selected-subtopics') {
            document.getElementById('finalize-subtopics-button').disabled = false;
        }
    }
}

function finalizeSelection() {
    document.getElementById('subtopic-selection').style.display = 'none';
    showCandidatePopup();
}

// Function to go back to the previous page
function goBack(previousPageId) {
    document.querySelectorAll('.start-container').forEach(container => {
        container.style.display = 'none'; // Hide all start containers
    });
    document.getElementById(previousPageId).style.display = 'block'; // Show the previous page
}

document.querySelectorAll('.selectBox').forEach(selectBox => {
    const checkboxes = selectBox.nextElementSibling;

    // When the selectBox is hovered, show the dropdown
    selectBox.addEventListener('mouseover', () => {
        checkboxes.style.display = 'block';
    });

    // When the cursor leaves both the selectBox and the checkboxes, hide the dropdown
    selectBox.addEventListener('mouseleave', () => {
        checkboxes.style.display = 'none';
    });

    checkboxes.addEventListener('mouseleave', () => {
        checkboxes.style.display = 'none';
    });

    // Keep the dropdown open when hovering over the checkboxes
    checkboxes.addEventListener('mouseover', () => {
        checkboxes.style.display = 'block';
    });

    // If clicking outside of the selectBox and checkboxes, do nothing (prevent focus loss issue)
    document.addEventListener('click', (event) => {
        if (!selectBox.contains(event.target) && !checkboxes.contains(event.target)) {
            // Do nothing, allowing the hover to re-trigger when hovering back over the selectBox
        }
    });
});

// Function to allow elements to be dragged over the drop zone
function allowDrop(event) {
    event.preventDefault();
}

// Function to handle the drag start event
function drag(event) {
    event.dataTransfer.setData('text', event.target.id);  // Set the id of the dragged element
}

// Function to handle the drop event
function drop(event) {
    event.preventDefault();
    const id = event.dataTransfer.getData('text');  // Get the id of the dragged element
    const draggedElement = document.getElementById(id);
    const selectedContainer = document.getElementById('selected-subtopics');

    if (!Array.from(selectedContainer.children).some(child => child.id === id)) {
        selectedContainer.appendChild(draggedElement);  // Move the element to the selected container
        draggedElement.classList.remove('subtopic-box'); // Remove the box from the top container
        draggedElement.classList.add('subtopic-box');  // Ensure it retains the correct styles in the bottom container
    }
    // Enable the Finalize button after at least one subtopic is added
    document.getElementById('finalize-subtopics-button').disabled = false;
}

// Function to populate subtopics on the third page
function populateSubtopics() {
    const availableSubtopicsContainer = document.getElementById('available-subtopics');
    availableSubtopicsContainer.innerHTML = ''; // Clear previous subtopics

    let positionMap = new Map(); // To keep track of the original positions
    let uniqueIdCounter = 0; // Counter to ensure unique IDs

    candidates.forEach(candidate => {
        selectedTopics.forEach(topic => {
            const subtopics = candidate.promises[topic]; // Get subtopics for the topic

            // Display each subtopic's promises
            Object.keys(subtopics).forEach(subtopic => {
                const promises = subtopics[subtopic];
                promises.forEach((promise, index) => {
                    // Skip subtopics with the description "Not addressed"
                    if (promise.description.trim() !== "Not addressed") {
                        const uniqueId = `promise-${topic}-${subtopic}-${index}-${uniqueIdCounter++}`; // Generate a unique ID
                        const subtopicBox = document.createElement('div');
                        subtopicBox.className = 'subtopic-box';
                        subtopicBox.id = uniqueId; // Assign the unique id
                        subtopicBox.draggable = true;
                        subtopicBox.ondragstart = drag;

                        // Store the original position
                        let originalPosition = positionMap.get(topic) || availableSubtopicsContainer.children.length;
                        subtopicBox.dataset.originalPosition = originalPosition;
                        positionMap.set(topic, originalPosition + 1);

                        subtopicBox.innerHTML = `
                            <p>${promise.description}</p>
                        `;
                        availableSubtopicsContainer.appendChild(subtopicBox);
                    }
                });
            });
        });
    });
}

function allowDrop(event) {
    event.preventDefault(); // Allow dropping by preventing the default handling
}

function drag(event) {
    const elementId = event.target.id;
    console.log(`Dragging element with ID: ${elementId}`); // Debugging: Log the ID of the element being dragged
    event.dataTransfer.setData('text/plain', elementId); // Store the id of the exact element being dragged
}

function drop(event) {
    event.preventDefault();
    const id = event.dataTransfer.getData('text'); // Retrieve the id of the dragged element
    console.log(`Dropping element with ID: ${id}`); // Debugging: Log the ID of the element being dropped
    const draggedElement = document.getElementById(id); // Get the exact element by its id

    const targetContainer = event.target.closest('.promise-list-container');

    // Ensure we are dropping in a valid container and not on the same container
    if (targetContainer && targetContainer !== draggedElement.parentElement) {
        // Move the exact dragged element to the target container
        targetContainer.appendChild(draggedElement);

        // If the element is moved to the selected subtopics container, add a remove button
        if (targetContainer.id === 'selected-subtopics') {
            if (!draggedElement.querySelector('button')) { // Check if the remove button already exists
                const removeButton = document.createElement('button');
                removeButton.textContent = 'Remove';
                removeButton.onclick = function() {
                    returnElementToOriginalPosition(draggedElement);
                };
                draggedElement.appendChild(removeButton); // Append the remove button to the dragged element
            }
        }

        updateFinalizeButton(); // Update finalize button state
    }
}

function returnElementToOriginalPosition(element) {
    const availableSubtopicsContainer = document.getElementById('available-subtopics');
    const originalPosition = parseInt(element.dataset.originalPosition, 10);
    const children = Array.from(availableSubtopicsContainer.children);

    // Insert the element back to its original position
    if (originalPosition >= children.length) {
        availableSubtopicsContainer.appendChild(element);
    } else {
        availableSubtopicsContainer.insertBefore(element, children[originalPosition]);
    }

    // Remove the remove button after moving back
    const removeButton = element.querySelector('button');
    if (removeButton) {
        element.removeChild(removeButton);
    }

    updateFinalizeButton(); // Update finalize button state
}

function updateFinalizeButton() {
    const selectedSubtopicsContainer = document.getElementById('selected-subtopics');
    document.getElementById('finalize-subtopics-button').disabled = selectedSubtopicsContainer.childElementCount === 0;
}

// Load the Excel file from the Google Sheets URL
loadExcelFromGoogleSheet('https://docs.google.com/spreadsheets/d/1_q6IaRErhJnPM_pTwiI7pGvOTJ4PJJRMTaz4zr-rmio/export?format=xlsx'); // Replace with the actual Google Sheets URL in export format