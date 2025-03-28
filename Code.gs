// Global constants
const DISPOSABLE_DOMAINS_URL = 'https://raw.githubusercontent.com/disposable/disposable-email-domains/master/domains.json';
const CACHE_DURATION = 21600; // 6 hours in seconds

/**
 * Custom function to validate and generate email addresses
 * @param {string} ownerName - The name of the person (optional)
 * @param {string} companyDomain - The company's domain (optional)
 * @param {string} existingEmail - Email address to validate (optional)
 * @return {string[][]} Validation results in a 2D array format for spreadsheet display
 * @customfunction
 */
function CUSTOM_EMAIL(ownerName, companyDomain, existingEmail) {
  // Initialize cache
  const cache = CacheService.getScriptCache();
  
  try {
    // If existing email is provided, validate it
    if (existingEmail) {
      const result = validateEmail(existingEmail);
      // Return as a 2D array with headers
      return [
        ['Email', 'Valid', 'Reason'],
        [result.email, result.valid.toString(), result.reason]
      ];
    }
    
    // If name and domain are provided, generate and validate email
    if (ownerName && companyDomain) {
      const generatedEmails = generateEmailVariations(ownerName, companyDomain);
      const results = validateEmailList(generatedEmails);
      
      // Convert results to 2D array with headers
      const output = [['Email', 'Valid', 'Reason']];
      results.forEach(result => {
        output.push([result.email, result.valid.toString(), result.reason]);
      });
      return output;
    }
    
    // Return error message if parameters are invalid
    return [
      ['Error'],
      ['Invalid parameters. Provide either an email to validate or name and domain to generate emails.']
    ];
  } catch (error) {
    // Return error message if something goes wrong
    return [
      ['Error'],
      [error.toString()]
    ];
  }
}

/**
 * Validates a single email address
 * @param {string} email - Email address to validate
 * @return {object} Validation results
 */
function validateEmail(email) {
  // Basic format validation
  const formatValid = validateEmailFormat(email);
  if (!formatValid) {
    return {
      email: email,
      valid: false,
      reason: 'Invalid email format'
    };
  }
  
  // Extract domain
  const domain = email.split('@')[1];
  
  // Check MX records
  const hasMX = checkMXRecord(domain);
  if (!hasMX) {
    return {
      email: email,
      valid: false,
      reason: 'Domain has no valid mail servers'
    };
  }
  
  // Check if disposable
  const isDisposable = checkDisposableDomain(domain);
  if (isDisposable) {
    return {
      email: email,
      valid: false,
      reason: 'Disposable email domain'
    };
  }
  
  return {
    email: email,
    valid: true,
    reason: 'Valid email address'
  };
}

/**
 * Validates email format using regex
 * @param {string} email - Email to validate
 * @return {boolean} Whether format is valid
 */
function validateEmailFormat(email) {
  const emailRegex = /^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)*$/;
  return emailRegex.test(email);
}

/**
 * Checks if domain has valid MX records
 * @param {string} domain - Domain to check
 * @return {boolean} Whether domain has MX records
 */
function checkMXRecord(domain) {
  try {
    const dnsQuery = UrlFetchApp.fetch(`https://dns.google/resolve?name=${domain}&type=MX`);
    const response = JSON.parse(dnsQuery.getContentText());
    return response.Answer && response.Answer.length > 0;
  } catch (e) {
    console.error('MX record check failed:', e);
    return false;
  }
}

/**
 * Checks if domain is a known disposable email domain
 * @param {string} domain - Domain to check
 * @return {boolean} Whether domain is disposable
 */
function checkDisposableDomain(domain) {
  const cache = CacheService.getScriptCache();
  let disposableDomains = cache.get('disposable_domains');
  
  if (!disposableDomains) {
    try {
      const response = UrlFetchApp.fetch(DISPOSABLE_DOMAINS_URL);
      disposableDomains = response.getContentText();
      cache.put('disposable_domains', disposableDomains, CACHE_DURATION);
    } catch (e) {
      console.error('Failed to fetch disposable domains:', e);
      return false;
    }
  }
  
  const domains = JSON.parse(disposableDomains);
  return domains.includes(domain.toLowerCase());
}

/**
 * Generates email variations from a name and domain
 * @param {string} name - Full name
 * @param {string} domain - Domain name
 * @return {string[]} List of possible email addresses
 */
function generateEmailVariations(name, domain) {
  const variations = [];
  const cleanDomain = domain.trim().toLowerCase();
  
  // Clean and split the name
  const nameParts = name.trim().toLowerCase().split(' ');
  const firstName = nameParts[0];
  const lastName = nameParts[nameParts.length - 1];
  
  // Generate variations
  variations.push(`${firstName}@${cleanDomain}`);
  variations.push(`${firstName}.${lastName}@${cleanDomain}`);
  variations.push(`${firstName[0]}${lastName}@${cleanDomain}`);
  variations.push(`${lastName}@${cleanDomain}`);
  variations.push(`${firstName}${lastName}@${cleanDomain}`);
  variations.push(`${firstName}_${lastName}@${cleanDomain}`);
  
  return variations;
}

/**
 * Validates a list of email addresses
 * @param {string[]} emails - List of emails to validate
 * @return {object[]} Validation results for each email
 */
function validateEmailList(emails) {
  const results = emails.map(email => validateEmail(email));
  // Sort by validity (valid emails first)
  return results.sort((a, b) => (b.valid ? 1 : 0) - (a.valid ? 1 : 0));
}

/**
 * Custom function to perform web searches using SERP API
 * @param {string} query - Search query
 * @param {number} maxResults - Maximum number of results (optional, default 5)
 * @return {string[][]} Search results in a 2D array format for spreadsheet display
 * @customfunction
 */
function CUSTOM_WEBSEARCH(query, maxResults = 5) {
  if (!query) {
    return [['Error'], ['Search query is required']];
  }

  const apiKey = getApiKey();
  if (!apiKey) {
    return [['Error'], ['SERP API key not set. Use "Data Enrichment > Set API Key" to configure.']];
  }

  try {
    const searchUrl = 'https://serpapi.com/search.json';
    const params = {
      q: query,
      api_key: apiKey,
      num: maxResults,
      engine: 'google'
    };
    
    const queryString = Object.keys(params)
      .map(key => `${encodeURIComponent(key)}=${encodeURIComponent(params[key])}`)
      .join('&');
    
    const response = UrlFetchApp.fetch(`${searchUrl}?${queryString}`);
    const data = JSON.parse(response.getContentText());
    
    // Format results for spreadsheet display
    const output = [['Title', 'Description', 'URL']];
    
    if (!data.organic_results || data.organic_results.length === 0) {
      return [['No Results'], ['No results found for the query'], ['']];
    }

    // Process only up to maxResults
    const results = data.organic_results.slice(0, maxResults);
    
    results.forEach(result => {
      output.push([
        result.title || 'No title',
        result.snippet || 'No description',
        result.link || 'No URL'
      ]);
    });

    return output;
  } catch (error) {
    console.error('Search error:', error);
    return [
      ['Error'],
      ['Failed to perform search. Error: ' + error.toString()],
      ['']
    ];
  }
}

/**
 * Shows a dialog to set the SERP API key
 */
function showSetApiKeyDialog() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    'Set SERP API Key',
    'Enter your SERP API key:',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() == ui.Button.OK) {
    const apiKey = result.getResponseText().trim();
    if (apiKey) {
      setApiKey(apiKey);
      ui.alert('API Key Saved', 'Your SERP API key has been saved successfully.', ui.ButtonSet.OK);
    } else {
      ui.alert('Error', 'API key cannot be empty.', ui.ButtonSet.OK);
    }
  }
}

/**
 * Gets the stored API key
 * @return {string} The stored API key or null if not set
 */
function getApiKey() {
  const scriptProperties = PropertiesService.getScriptProperties();
  return scriptProperties.getProperty('SERP_API_KEY');
}

/**
 * Sets the API key in script properties
 * @param {string} apiKey - The API key to store
 */
function setApiKey(apiKey) {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('SERP_API_KEY', apiKey);
}

/**
 * Test function for web search
 */
function testWebSearch() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  
  try {
    // Test search
    const testQuery = "artificial intelligence news";
    const result = CUSTOM_WEBSEARCH(testQuery, 3);
    
    // Write results to sheet
    const range = sheet.getRange(1, 5, result.length, 3); // Starting at column E
    range.setValues(result);
    
    ui.alert("Test completed! Check columns E-G for results.");
  } catch (error) {
    ui.alert("Error during test: " + error.toString());
  }
}

/**
 * Custom function to find people by job title at companies
 * @param {string} jobTitle - The job title to search for
 * @param {string} companyDomain - The company's domain
 * @param {number} maxResults - Maximum number of results (optional, default 3)
 * @return {string[][]} Search results in a 2D array format for spreadsheet display
 * @customfunction
 */
function CUSTOM_PERSON_LOOKUP(jobTitle, companyDomain, maxResults = 3) {
  if (!jobTitle || !companyDomain) {
    return [['Error'], ['Both job title and company domain are required']];
  }

  const apiKey = getApiKey();
  if (!apiKey) {
    return [['Error'], ['SERP API key not set. Use "Data Enrichment > Set API Key" to configure.']];
  }

  try {
    const searchUrl = 'https://serpapi.com/search.json';
    const query = `${jobTitle} at ${companyDomain} linkedin`;
    const params = {
      q: query,
      api_key: apiKey,
      num: maxResults,
      engine: 'google'
    };
    
    const queryString = Object.keys(params)
      .map(key => `${encodeURIComponent(key)}=${encodeURIComponent(params[key])}`)
      .join('&');
    
    const response = UrlFetchApp.fetch(`${searchUrl}?${queryString}`);
    const data = JSON.parse(response.getContentText());
    
    // Format results for spreadsheet display
    const output = [['Name', 'Title', 'Company', 'LinkedIn URL']];
    
    if (!data.organic_results || data.organic_results.length === 0) {
      return [['No Results'], ['No people found matching the criteria'], [''], ['']];
    }

    // Process only up to maxResults
    const results = data.organic_results.slice(0, maxResults);
    
    results.forEach(result => {
      // Extract name and details from the title
      const title = result.title || '';
      const name = title.split('|')[0].trim();
      const details = title.split('|')[1] || '';
      const company = details.split('-').pop().trim();
      
      output.push([
        name,
        jobTitle,
        company,
        result.link || 'No URL'
      ]);
    });

    return output;
  } catch (error) {
    console.error('Person lookup error:', error);
    return [
      ['Error'],
      ['Failed to perform person lookup. Error: ' + error.toString()],
      [''],
      ['']
    ];
  }
}

/**
 * Custom function to extract information from LinkedIn profiles
 * @param {string} linkedinUrl - The LinkedIn profile URL
 * @return {string[][]} Profile information in a 2D array format for spreadsheet display
 * @customfunction
 */
function CUSTOM_PERSON_LINKEDIN(linkedinUrl) {
  if (!linkedinUrl) {
    return [['Error'], ['LinkedIn URL is required']];
  }

  const apiKey = getApiKey();
  if (!apiKey) {
    return [['Error'], ['SERP API key not set. Use "Data Enrichment > Set API Key" to configure.']];
  }

  try {
    const searchUrl = 'https://serpapi.com/search.json';
    const params = {
      q: linkedinUrl,
      api_key: apiKey,
      num: 1,
      engine: 'google'
    };
    
    const queryString = Object.keys(params)
      .map(key => `${encodeURIComponent(key)}=${encodeURIComponent(params[key])}`)
      .join('&');
    
    const response = UrlFetchApp.fetch(`${searchUrl}?${queryString}`);
    const data = JSON.parse(response.getContentText());
    
    if (!data.organic_results || data.organic_results.length === 0) {
      return [['Error'], ['Could not find LinkedIn profile information']];
    }

    const result = data.organic_results[0];
    const snippet = result.snippet || '';
    
    // Extract information from the snippet
    const name = result.title ? result.title.split('-')[0].trim() : 'Unknown';
    const currentRole = snippet.match(/(?:Current:|at\s+)([^\.]+)/i)?.[1]?.trim() || 'Unknown';
    const location = snippet.match(/(?:Location:|based in\s+)([^\.]+)/i)?.[1]?.trim() || 'Unknown';
    
    // Return formatted profile information
    return [
      ['Field', 'Value'],
      ['Name', name],
      ['Current Role', currentRole],
      ['Location', location],
      ['Profile URL', linkedinUrl]
    ];
  } catch (error) {
    console.error('LinkedIn profile error:', error);
    return [
      ['Error'],
      ['Failed to extract LinkedIn profile information. Error: ' + error.toString()]
    ];
  }
}

// Add test functions for the new features
function testPersonLookup() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  
  try {
    // Test person lookup
    const result = CUSTOM_PERSON_LOOKUP('CEO', 'example.com', 2);
    
    // Write results to sheet
    const range = sheet.getRange(1, 9, result.length, 4); // Starting at column I
    range.setValues(result);
    
    ui.alert("Test completed! Check columns I-L for results.");
  } catch (error) {
    ui.alert("Error during test: " + error.toString());
  }
}

function testLinkedInProfile() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  
  try {
    // Test LinkedIn profile extraction
    const result = CUSTOM_PERSON_LINKEDIN('https://www.linkedin.com/in/example');
    
    // Write results to sheet
    const range = sheet.getRange(1, 13, result.length, 2); // Starting at column M
    range.setValues(result);
    
    ui.alert("Test completed! Check columns M-N for results.");
  } catch (error) {
    ui.alert("Error during test: " + error.toString());
  }
}

/**
 * Custom function to extract information from company LinkedIn profiles
 * @param {string} companyLinkedinUrl - The company's LinkedIn profile URL
 * @return {string[][]} Company information in a 2D array format for spreadsheet display
 * @customfunction
 */
function CUSTOM_BUSINESS_LINKEDIN(companyLinkedinUrl) {
  if (!companyLinkedinUrl) {
    return [['Error'], ['Company LinkedIn URL is required']];
  }

  const apiKey = getApiKey();
  if (!apiKey) {
    return [['Error'], ['SERP API key not set. Use "Data Enrichment > Set API Key" to configure.']];
  }

  try {
    const searchUrl = 'https://serpapi.com/search.json';
    const params = {
      q: companyLinkedinUrl,
      api_key: apiKey,
      num: 1,
      engine: 'google'
    };
    
    const queryString = Object.keys(params)
      .map(key => `${encodeURIComponent(key)}=${encodeURIComponent(params[key])}`)
      .join('&');
    
    const response = UrlFetchApp.fetch(`${searchUrl}?${queryString}`);
    const data = JSON.parse(response.getContentText());
    
    if (!data.organic_results || data.organic_results.length === 0) {
      return [['Error'], ['Could not find company LinkedIn profile information']];
    }

    const result = data.organic_results[0];
    const snippet = result.snippet || '';
    
    // Extract information from the snippet
    const name = result.title ? result.title.split('|')[0].trim() : 'Unknown';
    const industry = snippet.match(/(?:Industry:|in\s+)([^\.]+)/i)?.[1]?.trim() || 'Unknown';
    const size = snippet.match(/(?:Company size:|employees:?\s+)([^\.]+)/i)?.[1]?.trim() || 'Unknown';
    const location = snippet.match(/(?:Headquarters:|based in\s+)([^\.]+)/i)?.[1]?.trim() || 'Unknown';
    
    // Return formatted company information
    return [
      ['Field', 'Value'],
      ['Company Name', name],
      ['Industry', industry],
      ['Company Size', size],
      ['Location', location],
      ['LinkedIn URL', companyLinkedinUrl]
    ];
  } catch (error) {
    console.error('Company LinkedIn profile error:', error);
    return [
      ['Error'],
      ['Failed to extract company LinkedIn profile information. Error: ' + error.toString()]
    ];
  }
}

// Add test function for company LinkedIn profile
function testBusinessLinkedIn() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  
  try {
    // Test company LinkedIn profile extraction
    const result = CUSTOM_BUSINESS_LINKEDIN('https://www.linkedin.com/company/example');
    
    // Write results to sheet
    const range = sheet.getRange(1, 15, result.length, 2); // Starting at column O
    range.setValues(result);
    
    ui.alert("Test completed! Check columns O-P for results.");
  } catch (error) {
    ui.alert("Error during test: " + error.toString());
  }
}

/**
 * Custom function to find phone numbers by name and address
 * @param {string} personName - The person's full name
 * @param {string} address - The person's address or location
 * @return {string[][]} Phone number information in a 2D array format for spreadsheet display
 * @customfunction
 */
function CUSTOM_PHONE(personName, address) {
  if (!personName || !address) {
    return [['Error'], ['Both person name and address are required']];
  }

  const apiKey = getApiKey();
  if (!apiKey) {
    return [['Error'], ['SERP API key not set. Use "Data Enrichment > Set API Key" to configure.']];
  }

  try {
    const searchUrl = 'https://serpapi.com/search.json';
    const query = `${personName} ${address} phone number`;
    const params = {
      q: query,
      api_key: apiKey,
      num: 3,
      engine: 'google'
    };
    
    const queryString = Object.keys(params)
      .map(key => `${encodeURIComponent(key)}=${encodeURIComponent(params[key])}`)
      .join('&');
    
    const response = UrlFetchApp.fetch(`${searchUrl}?${queryString}`);
    const data = JSON.parse(response.getContentText());
    
    if (!data.organic_results || data.organic_results.length === 0) {
      return [['Error'], ['Could not find phone number information']];
    }

    // Extract phone numbers from search results
    const phoneNumbers = [];
    const phoneRegex = /(?:\+\d{1,3}[-.\s]?)?\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}/g;
    
    data.organic_results.forEach(result => {
      const snippet = result.snippet || '';
      const matches = snippet.match(phoneRegex);
      if (matches) {
        matches.forEach(number => {
          if (!phoneNumbers.includes(number)) {
            phoneNumbers.push(number);
          }
        });
      }
    });

    if (phoneNumbers.length === 0) {
      return [['No Results'], ['No phone numbers found for the given person']];
    }

    // Return formatted phone number information
    const output = [['Name', 'Address', 'Phone Number']];
    phoneNumbers.forEach(number => {
      output.push([personName, address, number]);
    });

    return output;
  } catch (error) {
    console.error('Phone lookup error:', error);
    return [
      ['Error'],
      ['Failed to find phone number. Error: ' + error.toString()]
    ];
  }
}

/**
 * Custom function to find phone numbers from LinkedIn profiles
 * @param {string} linkedinUrl - The LinkedIn profile URL
 * @return {string[][]} Phone number information in a 2D array format for spreadsheet display
 * @customfunction
 */
function CUSTOM_PHONE_LINKEDIN(linkedinUrl) {
  if (!linkedinUrl) {
    return [['Error'], ['LinkedIn URL is required']];
  }

  const apiKey = getApiKey();
  if (!apiKey) {
    return [['Error'], ['SERP API key not set. Use "Data Enrichment > Set API Key" to configure.']];
  }

  try {
    // First get the person's details from LinkedIn
    const profileInfo = CUSTOM_PERSON_LINKEDIN(linkedinUrl);
    if (profileInfo[0][0] === 'Error') {
      return profileInfo;
    }

    // Extract name and location from profile info
    const name = profileInfo.find(row => row[0] === 'Name')?.[1] || '';
    const location = profileInfo.find(row => row[0] === 'Location')?.[1] || '';

    if (!name || !location) {
      return [['Error'], ['Could not extract name and location from LinkedIn profile']];
    }

    // Use the extracted information to search for phone numbers
    return CUSTOM_PHONE(name, location);
  } catch (error) {
    console.error('LinkedIn phone lookup error:', error);
    return [
      ['Error'],
      ['Failed to find phone number from LinkedIn. Error: ' + error.toString()]
    ];
  }
}

// Add test function for phone number lookup
function testPhoneLookup() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  
  try {
    // Test phone number lookup
    const result = CUSTOM_PHONE('John Smith', 'New York, NY');
    
    // Write results to sheet
    const range = sheet.getRange(1, 17, result.length, 3); // Starting at column Q
    range.setValues(result);
    
    ui.alert("Test completed! Check columns Q-S for results.");
  } catch (error) {
    ui.alert("Error during test: " + error.toString());
  }
}

/**
 * Gets the stored Together API key
 * @return {string} The stored API key or null if not set
 */
function getTogetherApiKey() {
  const scriptProperties = PropertiesService.getScriptProperties();
  return scriptProperties.getProperty('TOGETHER_API_KEY');
}

/**
 * Sets the Together API key in script properties
 * @param {string} apiKey - The API key to store
 */
function setTogetherApiKey(apiKey) {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('TOGETHER_API_KEY', apiKey);
}

/**
 * Shows a dialog to set the Together API key
 */
function showSetTogetherApiKeyDialog() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    'Set Together API Key',
    'Enter your Together API key (get one for free at together.ai):\n\n',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() == ui.Button.OK) {
    try {
      const apiKey = result.getResponseText().trim();
      if (!apiKey) {
        throw new Error('API key cannot be empty');
      }
      setTogetherApiKey(apiKey);
      ui.alert(
        'API Key Saved',
        'Your Together API key has been saved successfully.',
        ui.ButtonSet.OK
      );
    } catch (error) {
      ui.alert('Error', error.message, ui.ButtonSet.OK);
    }
  }
}

/**
 * Custom function for AI-powered company analysis using Together.ai
 * @param {string} companyUrl - The company's website URL
 * @param {string} analysisInstructions - Instructions for analysis (e.g., "Return a 1-2 word industry description")
 * @return {string[][]} Analysis results in a 2D array format for spreadsheet display
 * @customfunction
 */
function CUSTOM_AGENT(companyUrl, analysisInstructions) {
  if (!companyUrl || !analysisInstructions) {
    return [['Error'], ['Both company URL and analysis instructions are required']];
  }

  const apiKey = getTogetherApiKey();
  if (!apiKey) {
    return [['Error'], ['Together API key not set. Use "Data Enrichment > Set Together API Key" to configure.']];
  }

  try {
    // First, get website summary using r.jina.ai
    const jinaUrl = `https://r.jina.ai/${encodeURIComponent(companyUrl)}`;
    const jinaResponse = UrlFetchApp.fetch(jinaUrl);
    const websiteContent = jinaResponse.getContentText();

    // Call Together.ai API for analysis
    const prompt = `Based on the following information about the company website ${companyUrl}:\n\n${websiteContent}\n\n${analysisInstructions}\n\nProvide a concise and accurate response.`;
    
    const togetherResponse = UrlFetchApp.fetch('https://api.together.xyz/v1/chat/completions', {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${apiKey}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        model: 'meta-llama/Llama-Vision-Free',
        messages: [
          {
            role: 'system',
            content: 'You are a business analyst AI that provides concise and accurate analysis of companies based on their website content.'
          },
          {
            role: 'user',
            content: prompt
          }
        ],
        temperature: 0.3,
        max_tokens: 500
      }),
      muteHttpExceptions: true
    });

    const aiData = JSON.parse(togetherResponse.getContentText());
    if (!aiData.choices || !aiData.choices[0] || !aiData.choices[0].message) {
      throw new Error('Invalid response from Together API');
    }

    const analysis = aiData.choices[0].message.content.trim();

    // Return formatted analysis
    return [
      ['Field', 'Value'],
      ['URL', companyUrl],
      ['Analysis Type', analysisInstructions],
      ['Result', analysis]
    ];
  } catch (error) {
    console.error('Company analysis error:', error);
    return [
      ['Error'],
      ['Failed to analyze company. Error: ' + error.toString()]
    ];
  }
}

/**
 * Shows the formula builder interface
 */
function showFormulaBuilder() {
  const html = HtmlService.createHtmlOutputFromFile('formulaBuilder')
    .setWidth(800)
    .setHeight(600)
    .setTitle('Custom Formula Builder');
  SpreadsheetApp.getUi().showModalDialog(html, 'Custom Formula Builder');
}

/**
 * Saves a custom formula definition
 */
function saveCustomFormula(formulaData) {
  // Validate formula data
  if (!formulaData.name || !formulaData.parameters || !formulaData.logic) {
    throw new Error('Missing required formula information');
  }

  // Format the formula name
  const formulaName = 'CUSTOM_' + formulaData.name.toUpperCase().replace(/[^A-Z0-9_]/g, '_');
  
  // Parse parameters
  const params = formulaData.parameters.split(',').map(p => p.trim());
  
  // Create the function definition
  const functionDef = {
    name: formulaName,
    parameters: params,
    logic: formulaData.logic,
    created: new Date().toISOString()
  };

  // Get existing formulas
  const userProperties = PropertiesService.getUserProperties();
  const formulas = JSON.parse(userProperties.getProperty('customFormulas') || '{}');
  
  // Add new formula
  formulas[formulaName] = functionDef;
  
  // Save updated formulas
  userProperties.setProperty('customFormulas', JSON.stringify(formulas));
  
  // Create the actual function
  _createCustomFunction(functionDef);
  
  return { name: formulaName };
}

/**
 * Tests a custom formula with sample input
 */
function testCustomFormula(testData) {
  if (!testData.logic) {
    throw new Error('Formula logic is required');
  }

  try {
    // Create a temporary function to test the logic
    const testFn = new Function(testData.testInput.split(',').map((_, i) => `arg${i}`).join(','), testData.logic);
    
    // Execute with test input
    const result = testFn.apply(null, testData.testInput.split(',').map(v => v.trim()));
    
    return result;
  } catch (error) {
    throw new Error('Test failed: ' + error.message);
  }
}

/**
 * Creates a custom function from a formula definition
 */
function _createCustomFunction(formulaDef) {
  try {
    // Create the function body
    const fnBody = `
      try {
        ${formulaDef.logic}
      } catch (error) {
        return error.message;
      }
    `;
    
    // Create the function
    const fn = new Function(formulaDef.parameters.join(','), fnBody);
    
    // Add it to the global scope
    this[formulaDef.name] = fn;
    
    return true;
  } catch (error) {
    throw new Error('Failed to create function: ' + error.message);
  }
}

/**
 * Loads all custom formulas when the spreadsheet opens
 */
function loadCustomFormulas() {
  const userProperties = PropertiesService.getUserProperties();
  const formulas = JSON.parse(userProperties.getProperty('customFormulas') || '{}');
  
  // Create all saved custom functions
  function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Data Enrichment')
      .addItem('Set SERP API Key', 'showSetApiKeyDialog')
      .addItem('Set Together API Key', 'showSetTogetherApiKeyDialog')
      .addSeparator()
      .addItem('Formula Builder', 'showFormulaBuilder')
      .addSeparator()
      .addItem('Run Email Test', 'testEmailValidation')
      .addItem('Run Web Search Test', 'testWebSearch')
      .addItem('Run Person Lookup Test', 'testPersonLookup')
      .addItem('Run LinkedIn Profile Test', 'testLinkedInProfile')
      .addItem('Run Company LinkedIn Test', 'testBusinessLinkedIn')
      .addItem('Run Phone Lookup Test', 'testPhoneLookup')
      .addItem('Run AI Analysis Test', 'testAgentAnalysis')
      .addSeparator()
      .addItem('About', 'showAbout')
      .addItem('Documentation', 'showDocs')
      .addToUi();
    
    // Load any saved custom formulas
    loadCustomFormulas();
  }

  /**
   * Shows the about dialog
   */
  function showAbout() {
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      'Data Enrichment Add-on',
      'This add-on provides custom functions for data enrichment and validation.\n\n' +
      'Available functions:\n' +
      '- CUSTOM_EMAIL(): Email validation and generation\n' +
      '- CUSTOM_WEBSEARCH(): Web search results\n' +
      '- CUSTOM_PERSON_LOOKUP(): Person lookup\n' +
      '- CUSTOM_PERSON_LINKEDIN(): LinkedIn profile extraction\n' +
      '- CUSTOM_BUSINESS_LINKEDIN(): Company LinkedIn profile extraction\n' +
      '- CUSTOM_PHONE(): Phone number lookup\n' +
      '- CUSTOM_PHONE_LINKEDIN(): LinkedIn phone number lookup\n' +
      '- CUSTOM_AGENT(): AI-powered company analysis using Together.ai\n' +
      '- More functions coming soon!',
      ui.ButtonSet.OK
    );
  }

  /**
   * Shows the documentation dialog
   */
  function showDocs() {
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      'Function Documentation',
      'CUSTOM_EMAIL(ownerName, companyDomain, existingEmail)\n' +
      '- ownerName: Person\'s full name (optional)\n' +
      '- companyDomain: Company\'s domain (optional)\n' +
      '- existingEmail: Email to validate (optional)\n\n' +
      'CUSTOM_WEBSEARCH(query, maxResults)\n' +
      '- query: Search term (required)\n' +
      '- maxResults: Maximum number of results (optional, default 5)\n\n' +
      'CUSTOM_PERSON_LOOKUP(jobTitle, companyDomain, maxResults)\n' +
      '- jobTitle: Job title to search for (required)\n' +
      '- companyDomain: Company\'s domain (required)\n' +
      '- maxResults: Maximum number of results (optional, default 3)\n\n' +
      'CUSTOM_PERSON_LINKEDIN(linkedinUrl)\n' +
      '- linkedinUrl: LinkedIn profile URL (required)\n\n' +
      'CUSTOM_BUSINESS_LINKEDIN(companyLinkedinUrl)\n' +
      '- companyLinkedinUrl: Company LinkedIn profile URL (required)\n\n' +
      'CUSTOM_PHONE(personName, address)\n' +
      '- personName: Person\'s full name (required)\n' +
      '- address: Person\'s address or location (required)\n\n' +
      'CUSTOM_PHONE_LINKEDIN(linkedinUrl)\n' +
      '- linkedinUrl: LinkedIn profile URL (required)\n\n' +
      'CUSTOM_AGENT(companyUrl, analysisInstructions)\n' +
      '- companyUrl: Company\'s website URL (required)\n' +
      '- analysisInstructions: Analysis instructions for the AI (required)\n' +
      '- Note: Uses r.jina.ai for website summarization and Together.ai for analysis\n\n' +
      'Examples:\n' +
      '=CUSTOM_EMAIL(,,user@example.com)\n' +
      '=CUSTOM_EMAIL("John Doe", "example.com")\n' +
      '=CUSTOM_WEBSEARCH("artificial intelligence")\n' +
      '=CUSTOM_WEBSEARCH("company news", 3)\n' +
      '=CUSTOM_PERSON_LOOKUP("CEO", "example.com")\n' +
      '=CUSTOM_PERSON_LINKEDIN("https://www.linkedin.com/in/example")\n' +
      '=CUSTOM_BUSINESS_LINKEDIN("https://www.linkedin.com/company/example")\n' +
      '=CUSTOM_PHONE("John Smith", "New York, NY")\n' +
      '=CUSTOM_PHONE_LINKEDIN("https://www.linkedin.com/in/example")\n' +
      '=CUSTOM_AGENT("https://example.com", "Analyze the company\'s industry and main offerings")',
      ui.ButtonSet.OK
    );
  }

  // Add a test function that can be run from the menu
  function testEmailValidation() {
    const ui = SpreadsheetApp.getUi();
    const sheet = SpreadsheetApp.getActiveSheet();
    
    try {
      // Test a valid email
      const testEmail = "test@gmail.com";
      const result = validateEmail(testEmail);
      
      // Write test results to the sheet
      sheet.getRange("A1").setValue("Test Results");
      sheet.getRange("A2").setValue("Email");
      sheet.getRange("B2").setValue("Valid");
      sheet.getRange("C2").setValue("Reason");
      
      sheet.getRange("A3").setValue(result.email);
      sheet.getRange("B3").setValue(result.valid.toString());
      sheet.getRange("C3").setValue(result.reason);
      
      ui.alert("Test completed! Check cells A1:C3 for results.");
    } catch (error) {
      ui.alert("Error during test: " + error.toString());
    }
  }

  // Add a test function that can be run from the menu
  function testAgentAnalysis() {
    const ui = SpreadsheetApp.getUi();
    const sheet = SpreadsheetApp.getActiveSheet();
    
    try {
      // Test a valid company analysis
      const testCompanyUrl = "https://example.com";
      const testAnalysisInstructions = "Analyze the company's industry and main offerings";
      const result = CUSTOM_AGENT(testCompanyUrl, testAnalysisInstructions);
      
      // Write test results to the sheet
      sheet.getRange("A1").setValue("Test Results");
      sheet.getRange("A2").setValue("URL");
      sheet.getRange("B2").setValue("Analysis Type");
      sheet.getRange("C2").setValue("Result");
      
      sheet.getRange("A3").setValue(result[0][0]);
      sheet.getRange("B3").setValue(result[0][1]);
      sheet.getRange("C3").setValue(result[0][2]);
      
      ui.alert("Test completed! Check cells A1:C3 for results.");
    } catch (error) {
      ui.alert("Error during test: " + error.toString());
    }
  }
}

/**
 * Handles API calls from the React application
 * @param {string} functionName - The name of the function to call
 * @param {Array} parameters - The parameters to pass to the function
 * @return {Object} The result of the function call
 */
function callFunction(functionName, parameters) {
  // Verify the function exists
  if (typeof this[functionName] !== 'function') {
    throw new Error(`Function ${functionName} not found`);
  }

  // Call the function with the provided parameters
  try {
    const result = this[functionName].apply(null, parameters);
    return { result };
  } catch (error) {
    throw new Error(`Error executing ${functionName}: ${error.message}`);
  }
}

/**
 * Custom function to perform web search and AI analysis on the results
 * @param {string} searchQuery - The search query
 * @param {string} formatInstructions - Instructions for formatting and analyzing the results
 * @param {number} maxResults - Maximum number of results (optional, default 3)
 * @return {string[][]} Formatted results in a 2D array for spreadsheet display
 * @customfunction
 */
function CUSTOM_SMART_SEARCH(searchQuery, formatInstructions, maxResults = 3) {
  if (!searchQuery) {
    return [['Error'], ['Search query is required']];
  }

  const serpApiKey = getApiKey();
  const togetherApiKey = getTogetherApiKey();
  
  if (!serpApiKey || !togetherApiKey) {
    return [['Error'], ['Both SERP API and Together API keys are required. Use "Data Enrichment > Set API Key" to configure.']];
  }

  try {
    // First, perform web search
    const searchResults = CUSTOM_WEBSEARCH(searchQuery, maxResults);
    if (searchResults[0][0] === 'Error' || searchResults[0][0] === 'No Results') {
      return searchResults;
    }

    // Remove header row and prepare content for AI analysis
    const contentToAnalyze = searchResults.slice(1).map(row => ({
      title: row[0],
      description: row[1],
      url: row[2]
    }));

    // Prepare prompt for AI analysis
    const prompt = `Based on these search results about "${searchQuery}":\n\n` +
      contentToAnalyze.map(result => 
        `Title: ${result.title}\nDescription: ${result.description}\nURL: ${result.url}\n`
      ).join('\n') +
      `\nPlease analyze and format the results according to these instructions: ${formatInstructions}`;

    // Call Together.ai API for analysis
    const togetherResponse = UrlFetchApp.fetch('https://api.together.xyz/v1/chat/completions', {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${togetherApiKey}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        model: 'meta-llama/Llama-Vision-Free',
        messages: [
          {
            role: 'system',
            content: 'You are an AI analyst that provides insightful analysis of web search results in the requested format.'
          },
          {
            role: 'user',
            content: prompt
          }
        ],
        temperature: 0.3,
        max_tokens: 1000
      }),
      muteHttpExceptions: true
    });

    const aiData = JSON.parse(togetherResponse.getContentText());
    if (!aiData.choices || !aiData.choices[0] || !aiData.choices[0].message) {
      throw new Error('Invalid response from Together API');
    }

    const analysis = aiData.choices[0].message.content.trim();

    // Return formatted results
    return [
      ['Search Query', 'AI Analysis'],
      [searchQuery, analysis]
    ];

  } catch (error) {
    console.error('Smart search error:', error);
    return [
      ['Error'],
      ['Failed to perform smart search. Error: ' + error.toString()]
    ];
  }
}

// Add test function for smart search
function testSmartSearch() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  
  try {
    // Test smart search
    const result = CUSTOM_SMART_SEARCH(
      "AI developments in Indonesia last week",
      "Format each result as: [Title], [Brief Summary], [URL]. Then add a paragraph analyzing the trends."
    );
    
    // Write results to sheet
    const range = sheet.getRange(1, 19, result.length, 2); // Starting at column S
    range.setValues(result);
    
    ui.alert("Test completed! Check columns S-T for results.");
  } catch (error) {
    ui.alert("Error during test: " + error.toString());
  }
} 