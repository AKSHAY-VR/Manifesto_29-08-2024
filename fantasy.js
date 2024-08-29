let candidates = [];
let topics = [];
let selectedTopics = [];
let selectedPromiseCategories = [];
let selectedCandidates = [];

function loadExcelFromGoogleSheet(sheetUrl) {
    fetch(sheetUrl)
        .then(response => response.arrayBuffer())
        .then(data => {
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
            console.log("Excel Data:", jsonData);
            processExcelData(jsonData);
        })
        .catch(error => console.error('Error loading Excel file from Google Sheets:', error));
}

function processExcelData(data) {
    const topicsData = {};

    data.forEach(row => {
        const topic = row['Topic'];
        const promiseCategory = row['Promise Categories'];
        const candidate = row['Candidate'];
        const candidatePromise = row['Candidate promise'] || '';

        if (!topicsData[topic]) {
            topicsData[topic] = {};
        }

        if (!topicsData[topic][promiseCategory]) {
            topicsData[topic][promiseCategory] = {};
        }

        if (!topicsData[topic][promiseCategory][candidate]) {
            topicsData[topic][promiseCategory][candidate] = [];
        }

        if (candidatePromise) {
            topicsData[topic][promiseCategory][candidate].push(candidatePromise);
        }
    });

    candidates = [
        { name: 'Ranil Wickremesinghe', image: 'candidate1.png', promises: topicsData, totalPromises: 0 },
        { name: 'Anura Dissanayake', image: 'candidate2.png', promises: topicsData, totalPromises: 0 },
        { name: 'Namal Rajapaksa', image: 'candidate3.png', promises: topicsData, totalPromises: 0 },
    ];

    topics = Object.keys(topicsData);
    console.log("Processed Topics Data:", topicsData);
    populateTopicOptions();
}

function populateCandidatePromises() {
    const candidatePromisesContainer = document.getElementById('candidate-options');
    candidatePromisesContainer.innerHTML = ''; // Clear any existing content

    let hasPromises = false;

    // Iterate over each selected topic
    selectedTopics.forEach(topic => {
        // Iterate over each selected promise category
        selectedPromiseCategories.forEach(category => {
            // For each candidate, display the promises related to the selected topic and category
            candidates.forEach(candidate => {
                const promises = candidate.promises[topic] && candidate.promises[topic][category] && candidate.promises[topic][category][candidate.name];

                if (promises && promises.length > 0) {
                    hasPromises = true;
                    const candidateCard = document.createElement('div');
                    candidateCard.className = 'candidate-card';
                    candidateCard.innerHTML = `
                        <h4>${candidate.name} - ${category} (${topic})</h4>
                        <ul>
                            ${promises.map(promise => `<li>${promise}</li>`).join('')}
                        </ul>
                    `;
                    candidatePromisesContainer.appendChild(candidateCard);
                }
            });
        });
    });

    if (!hasPromises) {
        candidatePromisesContainer.innerHTML = '<p>No promises found for the selected categories.</p>';
    }
}

function startFantasyGame() {
    showContainer('topic-selection');
}

function populateTopicOptions() {
    const topicOptionsContainer = document.getElementById('topic-options');
    topicOptionsContainer.innerHTML = '';

    topics.forEach(topic => {
        const label = document.createElement('label');
        label.innerHTML = `<input type="checkbox" name="topic" value="${topic}" onchange="updateSelectedTopics()"> ${topic}`;
        topicOptionsContainer.appendChild(label);
    });
}

function updateSelectedTopics() {
    selectedTopics = Array.from(document.querySelectorAll('#topic-options input:checked')).map(input => input.value);
    document.getElementById('confirm-topic-button').disabled = selectedTopics.length === 0;
}

function confirmTopic() {
    showContainer('promise-category-selection');
    populatePromiseCategoryOptions();
}

function populatePromiseCategoryOptions() {
    const promiseCategoryOptionsContainer = document.getElementById('promise-category-options');
    promiseCategoryOptionsContainer.innerHTML = '';

    let uniquePromiseCategories = new Set();
    selectedTopics.forEach(topic => {
        const promiseCategories = Object.keys(candidates[0].promises[topic]);
        promiseCategories.forEach(category => uniquePromiseCategories.add(category));
    });

    uniquePromiseCategories = Array.from(uniquePromiseCategories);

    uniquePromiseCategories.forEach(category => {
        const label = document.createElement('label');
        label.innerHTML = `<input type="checkbox" name="promise-category" value="${category}" onchange="updateSelectedPromiseCategories()"> ${category}`;
        promiseCategoryOptionsContainer.appendChild(label);
    });
}

function updateSelectedPromiseCategories() {
    selectedPromiseCategories = Array.from(document.querySelectorAll('#promise-category-options input:checked')).map(input => input.value);
    document.getElementById('confirm-promise-category-button').disabled = selectedPromiseCategories.length === 0;
}

function confirmPromiseCategory() {
    showContainer('candidate-selection');
    populateCandidatePromises(); // Call this to populate the fourth page
}

function updateSelectedCandidates() {
    selectedCandidates = Array.from(document.querySelectorAll('#candidate-options input:checked')).map(input => input.value);
    document.getElementById('finalize-candidate-button').disabled = selectedCandidates.length === 0;
}

function finalizeSelection() {
    showContainer('result-page');
    showCandidatePopup();
}

function showCandidatePopup() {
    const candidateCardsContainer = document.getElementById('top-candidates');
    candidateCardsContainer.innerHTML = '';

    let totalPromises = 0;

    candidates.forEach(candidate => {
        selectedTopics.forEach(topic => {
            selectedPromiseCategories.forEach(category => {
                if (candidate.promises[topic] && candidate.promises[topic][category] && candidate.promises[topic][category][candidate.name]) {
                    candidate.totalPromises += candidate.promises[topic][category][candidate.name].length;
                    totalPromises += candidate.promises[topic][category][candidate.name].length;
                }
            });
        });
    });

    candidates.forEach(candidate => {
        const percentage = (candidate.totalPromises / totalPromises) * 100 || 0;
        const card = document.createElement('div');
        card.className = 'candidate-card';
        card.innerHTML = `
            <img src="${candidate.image}" alt="${candidate.name}">
            <h3>${candidate.name}</h3>
            <p>${candidate.totalPromises} promises selected (${percentage.toFixed(2)}%)</p>
        `;
        candidateCardsContainer.appendChild(card);
    });
}

function showContainer(containerId) {
    document.querySelectorAll('.start-container').forEach(container => {
        container.style.display = 'none';
    });
    document.getElementById(containerId).style.display = 'block';
}

function goBack(previousPageId) {
    showContainer(previousPageId);
}

function goToManifesto() {
    window.location.href = 'index.html';
}

loadExcelFromGoogleSheet('https://docs.google.com/spreadsheets/d/15WNSsgyRZ295kpDu58kbCWJrkZ2SxnNi6ce7Wjog43s/export?format=xlsx');
