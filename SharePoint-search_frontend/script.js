// Function to search SharePoint based on input query
async function searchSharePoint() {
  const query = document.getElementById('search-query').value;

  if (!query) {
      alert('Please enter a search term.');
      return;
  }

  // Disable the search button while the search is in progress
  const searchButton = document.querySelector('button');
  searchButton.disabled = true;

  try {
      const results = await fetchSearchResults(query);
      displaySearchResults(results);
  } catch (error) {
      alert('Error occurred while searching.');
      console.error(error);
  } finally {
      // Enable the search button again
      searchButton.disabled = false;
  }
}

// Function to fetch search results from the backend API
async function fetchSearchResults(query) {
  const response = await fetch('https://localhost:7064/api/SharePointSearch/search', {
      method: 'POST',
      headers: {
          'Content-Type': 'application/json'
      },
      body: JSON.stringify({ query: query })
  });

  if (!response.ok) {
      throw new Error('Failed to fetch search results');
  }

  const data = await response.json();
  return data;
}

// Function to display search results on the page
function displaySearchResults(results) {
  const resultsList = document.getElementById('search-results-list');
  resultsList.innerHTML = ''; // Clear previous results

  if (results.length === 0) {
      resultsList.innerHTML = '<li>No results found.</li>';
      return;
  }

  results.forEach(result => {
      const li = document.createElement('li');
      li.innerHTML = `
          <h3><a href="${result.webUrl}" target="_blank">${result.name}</a></h3>
          <p><strong>Last Modified By:</strong> ${result.lastModifiedBy}</p>
          <p><strong>Web URL:</strong> <a href="${result.webUrl}" target="_blank">${result.webUrl}</a></p>
      `;
      resultsList.appendChild(li);
  });
}
