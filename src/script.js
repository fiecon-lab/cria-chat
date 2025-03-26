/**
 * Cria Chat - A markdown-enabled chat interface
 * This script handles the chat functionality, markdown parsing, and API integration
 */

// ===== LOAD DEPENDENCIES =====

// Dynamically load JSZip library
function loadJSZip() {
  return new Promise((resolve, reject) => {
    if (typeof JSZip !== 'undefined') {
      resolve();
      return;
    }
    
    const script = document.createElement('script');
    script.src = 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js';
    script.integrity = 'sha512-XMVd28F1oH/O71fzwBnV7HucLxVwtxf26XV8P4wPk26EDxuGZ91N8bsOttmnomcCD3CS5ZMRL50H0GgOHvegtg==';
    script.crossOrigin = 'anonymous';
    script.onload = () => resolve();
    script.onerror = () => reject(new Error('Failed to load JSZip'));
    document.head.appendChild(script);
  });
}

// Dynamically load xmldom library
function loadXMLDOM() {
  return new Promise((resolve, reject) => {
    // If DOMParser is already defined in the browser, use it
    if (typeof DOMParser !== 'undefined') {
      resolve();
      return;
    }
    
    // Try to load the @xmldom/xmldom library
    const script = document.createElement('script');
    script.src = 'https://cdn.jsdelivr.net/npm/@xmldom/xmldom@0.9.8/lib/index.min.js';
    script.onload = () => {
      // Check if the library exposed xmldom or @xmldom/xmldom
      if (typeof window.xmldom !== 'undefined' && window.xmldom.DOMParser) {
        // Expose DOMParser globally
        window.DOMParser = window.xmldom.DOMParser;
        resolve();
      } else if (typeof window.DOMParser !== 'undefined') {
        // DOMParser is already exposed
        resolve();
      } else {
        // Try to find DOMParser in other possible locations
        console.warn("xmldom loaded but DOMParser not found directly. Trying alternative approaches.");
        if (typeof window['@xmldom/xmldom'] !== 'undefined' && window['@xmldom/xmldom'].DOMParser) {
          window.DOMParser = window['@xmldom/xmldom'].DOMParser;
          resolve();
        } else {
          // Try loading fast-xml-parser as a fallback
          console.warn("Failed to find DOMParser in loaded xmldom library. Trying fast-xml-parser as fallback.");
          loadFastXMLParser()
            .then(resolve)
            .catch(reject);
        }
      }
    };
    script.onerror = () => {
      console.warn("Failed to load xmldom. Trying fast-xml-parser as fallback.");
      loadFastXMLParser()
        .then(resolve)
        .catch(reject);
    };
    document.head.appendChild(script);
  });
}

// Dynamically load fast-xml-parser as a fallback
function loadFastXMLParser() {
  return new Promise((resolve, reject) => {
    const script = document.createElement('script');
    script.src = 'https://cdnjs.cloudflare.com/ajax/libs/fast-xml-parser/4.5.1/fxparser.min.js';
    script.onload = () => {
      // Create a DOMParser-like interface using fast-xml-parser
      if (typeof window.fxparser !== 'undefined' && window.fxparser.XMLParser) {
        console.log("Using fast-xml-parser as DOMParser fallback");
        
        // Create a DOMParser-like class that uses fast-xml-parser internally
        window.DOMParser = class FXPDOMParser {
          parseFromString(xmlString, mimeType) {
            const parser = new window.fxparser.XMLParser({
              ignoreAttributes: false,
              attributeNamePrefix: "",
              preserveOrder: true
            });
            
            // Parse XML to JS object
            const jsObj = parser.parse(xmlString);
            
            // Create a simple document-like object with getElementsByTagName
            const doc = {
              documentElement: jsObj[0],
              getElementsByTagName: function(tagName) {
                const result = [];
                
                // Simple recursive function to find elements by tag name
                function findElements(obj, tagName) {
                  if (Array.isArray(obj)) {
                    for (const item of obj) {
                      findElements(item, tagName);
                    }
                  } else if (obj && typeof obj === 'object') {
                    // Check if this object is an element with the matching tag name
                    if (obj.tagName && obj.tagName === tagName) {
                      result.push({
                        textContent: obj.value || "",
                        // Add other properties as needed
                        getElementsByTagName: function(childTagName) {
                          const childResult = [];
                          if (obj.children) {
                            for (const child of obj.children) {
                              if (child.tagName === childTagName) {
                                childResult.push({
                                  textContent: child.value || ""
                                });
                              }
                            }
                          }
                          return {
                            length: childResult.length,
                            item: function(index) {
                              return childResult[index];
                            }
                          };
                        }
                      });
                    }
                    
                    // Recursively search children
                    if (obj.children) {
                      findElements(obj.children, tagName);
                    }
                  }
                }
                
                findElements(jsObj, tagName);
                
                // Return a NodeList-like object
                return {
                  length: result.length,
                  item: function(index) {
                    return result[index];
                  }
                };
              }
            };
            
            return doc;
          }
        };
        
        resolve();
      } else {
        reject(new Error('Failed to find XMLParser in loaded fast-xml-parser library'));
      }
    };
    script.onerror = () => reject(new Error('Failed to load fast-xml-parser'));
    document.head.appendChild(script);
  });
}

// Load dependencies
Promise.all([loadJSZip(), loadXMLDOM()])
  .then(() => console.log('Dependencies loaded successfully'))
  .catch(error => {
    console.error('Error loading dependencies:', error);
    // Log more specific error information
    if (error.message.includes('xmldom') && error.message.includes('fast-xml-parser')) {
      console.error('Both xmldom and fast-xml-parser failed to load. XML parsing functionality may be limited.');
    }
  });

// ===== CONFIGURATION =====

// Configure marked options to work with Prism.js
marked.setOptions({
  highlight: function(code, lang) {
    // Use a try-catch to handle any highlighting errors
    try {
      if (lang) {
        // Convert language name to what Prism expects
        const language = getPrismLanguage(lang);
        if (Prism.languages[language]) {
          return Prism.highlight(code, Prism.languages[language], language);
        }
      }
      // Default to javascript if no language is specified
      return Prism.highlight(code, Prism.languages.javascript, 'javascript');
    } catch (e) {
      console.error("Prism error:", e);
      return code; // Return unhighlighted code if there's an error
    }
  },
  langPrefix: 'language-',
  breaks: true,
  gfm: true
});

// ===== CONSTANTS & VARIABLES =====

// Store conversation history
let conversationHistory = [];

// Office application type
let officeAppType = null;

// DOM elements
const chatMessages = document.getElementById('chat-messages');
const messageInput = document.getElementById('message-input');
const sendButton = document.getElementById('send-button');
const typingIndicator = document.getElementById('typing-indicator');
const newChatButtonDesktop = document.getElementById('new-chat-button-desktop');
const newChatButtonMobile = document.getElementById('new-chat-button-mobile');
const docContentCheckbox = document.getElementById('doc-content');
const docContentLabel = document.querySelector('label[for="doc-content"]');
const docAccessHint = document.getElementById('doc-access-hint');

// API details
const API_URL = "https://cria-api.fiecon.com/api/generate";
const API_KEY = typeof config !== 'undefined' ? config.API_KEY : '';
const API_MODEL = "llama3.2-vision:latest";

// Add error handling for missing API key
if (!API_KEY) {
  console.error('API key not found. Please ensure config.js is loaded properly.');
}

// Welcome message text
const WELCOME_MESSAGE = "How can I assist you today?";

// ===== LANGUAGE UTILITIES =====

/**
 * Converts language shorthand to the format Prism expects
 * @param {string} lang - The language identifier
 * @returns {string} - The Prism-compatible language name
 */
function getPrismLanguage(lang) {
  const langMap = {
    'js': 'javascript',
    'py': 'python',
    'ts': 'typescript',
    'html': 'markup',
    'xml': 'markup',
    'sh': 'bash',
    'shell': 'bash',
    'cs': 'csharp'
  };
  
  return langMap[lang.toLowerCase()] || lang.toLowerCase();
}

/**
 * Returns a display-friendly name for a programming language
 * @param {string} language - The language identifier
 * @returns {string} - User-friendly language name
 */
function getLanguageDisplayName(language) {
  const languageMap = {
    'js': 'JavaScript',
    'javascript': 'JavaScript',
    'ts': 'TypeScript',
    'typescript': 'TypeScript',
    'py': 'Python',
    'python': 'Python',
    'java': 'Java',
    'cs': 'C#',
    'csharp': 'C#',
    'html': 'HTML',
    'css': 'CSS',
    'json': 'JSON',
    'xml': 'XML',
    'sql': 'SQL',
    'bash': 'Bash',
    'sh': 'Shell',
    'shell': 'Shell',
    'plaintext': 'Plain Text',
    'text': 'Plain Text',
    'markup': 'HTML/XML',
    'code': 'Code'
  };
  
  return languageMap[language.toLowerCase()] || language.charAt(0).toUpperCase() + language.slice(1);
}

/**
 * Generates a dynamic system prompt based on the current date and Office application context
 * @returns {string} - The formatted system prompt
 */
function generateSystemPrompt() {
  // Get current date in a readable format
  const now = new Date();
  const currentDateTime = now.toLocaleDateString('en-US', { 
    year: 'numeric', 
    month: 'long', 
    day: 'numeric' 
  });
  
  // Base prompt template
  let prompt = `The assistant is Cria, created by FIECON, a real consultancy specializing in health economics and market access. The current date is ${currentDateTime}. Cria's knowledge base was last updated in December 2023 and it answers user questions about events before December 2023 and after December 2023 the same way a highly informed health economics professional from December 2023 would if they were talking to someone from ${currentDateTime}. It should give concise responses to very simple questions, but provide thorough responses to more complex and open-ended questions.`;
  
  // Add context-specific information based on Office application
  if (officeAppType) {
    if (officeAppType === Office.HostType.Word) {
      prompt += ` The user is currently working in Microsoft Word. It is happy to help with document writing, formatting, and content creation.`;
    } else if (officeAppType === Office.HostType.PowerPoint) {
      prompt += ` The user is currently working in Microsoft PowerPoint. It is happy to help with presentation content, slide design, and creating compelling narratives.`;
    } else if (officeAppType === Office.HostType.Excel) {
      prompt += ` The user is currently working in Microsoft Excel. It is happy to help with data analysis, formulas, and VBA code.`;
    } 
  } else {
    prompt += ` It is happy to help with health economics analyses, healthcare value assessments, market access strategies, and health technology assessments.`;
  }
  
  // Add final part of the prompt
  prompt += ` It uses markdown for coding. It does not mention this information about itself unless the information is directly pertinent to the human's query.`;
  
  return prompt;
}

// ===== EVENT LISTENERS =====

// Initialize with a welcome message and Office.js
window.onload = function() {
  // Add welcome message
  addBotMessage(WELCOME_MESSAGE);
  
  // Make sure Prism highlights all code blocks
  setTimeout(() => {
    Prism.highlightAll();
  }, 100);
  
  // Make textarea auto-resize on load and hide scrollbar
  messageInput.style.height = 'auto';
  messageInput.style.overflowY = 'hidden';
  
  // Focus on input field
  messageInput.focus();
  
  // Initialize Office.js
  if (typeof Office !== 'undefined') {
    Office.onReady(function(info) {
      officeAppType = info.host;
      
      // Configure the document content checkbox based on the Office application
      configureDocContentCheckbox();
      
      // Add debug button for PowerPoint
      if (officeAppType === Office.HostType.PowerPoint) {
        // Create a debug button
        const debugButton = document.createElement('button');
        debugButton.textContent = 'Debug Slide Files';
        debugButton.className = 'btn btn-sm btn-secondary mt-2';
        debugButton.style.display = 'none'; // Hide by default
        debugButton.onclick = debugListPresentationFiles;
        
        // Add button to the UI
        const chatControls = document.querySelector('.chat-controls');
        if (chatControls) {
          chatControls.appendChild(debugButton);
        }
        
        // Show debug button with Ctrl+Shift+D
        document.addEventListener('keydown', function(e) {
          if (e.ctrlKey && e.shiftKey && e.key === 'D') {
            debugButton.style.display = debugButton.style.display === 'none' ? 'block' : 'none';
          }
        });
      }
    });
  } else {
    console.warn("Office.js is not available. Running in standalone mode.");
    // Hide the document content checkbox if not in an Office context
    if (docContentCheckbox && docContentCheckbox.parentElement) {
      docContentCheckbox.parentElement.style.display = 'none';
    }
  }
};

// Configure the document content checkbox based on the Office application
function configureDocContentCheckbox() {
  if (!docContentCheckbox) return;
  
  if (officeAppType === Office.HostType.Excel) {
    // Disable the checkbox in Excel
    docContentCheckbox.disabled = true;
    docContentCheckbox.checked = false;
    if (docContentLabel) {
      docContentLabel.style.color = '#999';
      docContentLabel.textContent = 'Document content not yet available in Excel';
    }
    if (docAccessHint) {
      docAccessHint.textContent = 'Document content access is not yet available in Excel.';
    }
  } else if (officeAppType === Office.HostType.Word) {
    // Enable the checkbox in Word
    docContentCheckbox.disabled = false;
    if (docContentLabel) {
      docContentLabel.style.color = '';
      docContentLabel.textContent = 'Allow access to the document';
    }
    if (docAccessHint) {
      docAccessHint.textContent = 'Check "Allow access to the document" to include document content in your conversation.';
    }
  } else if (officeAppType === Office.HostType.PowerPoint) {
    // Enable the checkbox in PowerPoint
    docContentCheckbox.disabled = false;
    if (docContentLabel) {
      docContentLabel.style.color = '';
      docContentLabel.textContent = 'Allow access to the active slide';
    }
    if (docAccessHint) {
      docAccessHint.textContent = 'Check to include slide content. For best results, select specific text on the slide.';
    }
  } else {
    // Hide the checkbox for other contexts
    if (docContentCheckbox.parentElement) {
      docContentCheckbox.parentElement.style.display = 'none';
    }
    if (docAccessHint) {
      docAccessHint.textContent = 'Cria cannot see or directly access your document.';
    }
  }
}

// Auto-resize textarea as user types
messageInput.addEventListener('input', function() {
  this.style.height = 'auto';
  this.style.height = (this.scrollHeight) + 'px';
  
  // Limit max height
  const maxHeight = window.innerHeight * 0.3; // 30% of viewport height
  if (this.scrollHeight > maxHeight) {
    this.style.height = maxHeight + 'px';
    this.style.overflowY = 'auto'; // Show scrollbar only when content exceeds max height
  } else {
    this.style.overflowY = 'hidden'; // Hide scrollbar when not needed
  }
});

// Reset textarea height when cleared
messageInput.addEventListener('keydown', function(e) {
  if (e.key === 'Enter' && !e.shiftKey) {
    // Will be cleared after sending, prepare for reset
    setTimeout(() => {
      this.style.height = 'auto';
      this.style.overflowY = 'hidden';
    }, 0);
  }
});

// Add responsive behavior
window.addEventListener('resize', function() {
  // Adjust textarea max height on window resize
  const maxHeight = window.innerHeight * 0.3;
  if (messageInput.scrollHeight > maxHeight) {
    messageInput.style.height = maxHeight + 'px';
  }
});

// Send message on button click
sendButton.addEventListener('click', sendMessage);

// Send message on Enter key (without Shift)
messageInput.addEventListener('keypress', function(e) {
  if (e.key === 'Enter' && !e.shiftKey) {
    e.preventDefault();
    sendMessage();
  }
});

// Start new conversation on button click
newChatButtonDesktop.addEventListener('click', startNewConversation);
newChatButtonMobile.addEventListener('click', startNewConversation);

// Add event listener for the document content checkbox
docContentCheckbox.addEventListener('change', function() {
  console.log(`Document content access: ${this.checked ? 'enabled' : 'disabled'}`);
  
  // Update the hint text based on checkbox state
  if (docAccessHint) {
    if (this.checked && !this.disabled) {
      if (officeAppType === Office.HostType.PowerPoint) {
        docAccessHint.textContent = 'Slide content will be included in your conversation.';
      } else {
        docAccessHint.textContent = 'Document content will be included in your conversation.';
      }
    } else if (!this.disabled) {
      if (officeAppType === Office.HostType.PowerPoint) {
        docAccessHint.textContent = 'Check to include slide content. For best results, select specific text on the slide.';
      } else {
        docAccessHint.textContent = 'Check "Allow access to the document" to include document content in your conversation.';
      }
    }
  }
});

// ===== CORE FUNCTIONALITY =====

/**
 * Starts a new conversation by clearing history and displaying welcome message
 */
function startNewConversation() {
  // Clear chat history
  conversationHistory = [];
  
  // Clear chat messages
  chatMessages.innerHTML = '';
  
  // Add welcome message
  addBotMessage(WELCOME_MESSAGE);
  
  // Make sure Prism highlights all code blocks
  setTimeout(() => {
    Prism.highlightAll();
  }, 100);
  
  // Reset textarea and hide scrollbar
  messageInput.value = '';
  messageInput.style.height = 'auto';
  messageInput.style.overflowY = 'hidden';
  
  // Reset document content checkbox if it's not disabled
  if (docContentCheckbox && !docContentCheckbox.disabled) {
    docContentCheckbox.checked = false;
    
    // Reset the hint text
    if (docAccessHint) {
      if (officeAppType === Office.HostType.PowerPoint) {
        docAccessHint.textContent = 'Check to include slide content. For best results, select specific text on the slide.';
      } else {
        docAccessHint.textContent = 'Check "Allow access to the document" to include document content in your conversation.';
      }
    }
  }
  
  // Focus on input field
  messageInput.focus();
}

/**
 * Sends user message to the API and handles the response
 */
async function sendMessage() {
  const userMessage = messageInput.value.trim();
  if (!userMessage) return;
  
  // Clear input field and reset height
  messageInput.value = '';
  messageInput.style.height = 'auto';
  messageInput.style.overflowY = 'hidden'; // Ensure scrollbar is hidden after sending
  
  // Add user message to chat
  addUserMessage(userMessage);
  
  // Update conversation history
  conversationHistory.push({ role: "user", content: userMessage });
  
  // Show typing indicator
  typingIndicator.style.display = 'flex';
  
  try {
    // Check if document content should be included
    let documentContent = "";
    if (docContentCheckbox && docContentCheckbox.checked && !docContentCheckbox.disabled) {
      documentContent = await getDocumentContent();
    }
    
    // Prepare the prompt with conversation history
    const prompt = formatConversationForAPI(conversationHistory, documentContent);
    
    // Send request to API
    const response = await fetch(API_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        APIKey: API_KEY,
      },
      body: JSON.stringify({
        model: API_MODEL,
        prompt: prompt,
        stream: false,
      }),
    });
    
    if (!response.ok) {
      throw new Error(`HTTP error! Status: ${response.status}`);
    }
    
    // Parse the JSON response
    const data = await response.json();
    
    // Hide typing indicator
    typingIndicator.style.display = 'none';
    
    // Add bot response to chat
    const botResponse = data.response;
    addBotMessage(botResponse);
    
    // Update conversation history
    conversationHistory.push({ role: "assistant", content: botResponse });
    
  } catch (error) {
    console.error("Error:", error);
    typingIndicator.style.display = 'none';
    addBotMessage(`Error: ${error.message}`);
  }
}

/**
 * Gets the content of the current document
 * @returns {Promise<string>} - The document content
 */
async function getDocumentContent() {
  return new Promise((resolve, reject) => {
    try {
      if (!Office || !officeAppType) {
        resolve("");
        return;
      }
      
      if (officeAppType === Office.HostType.Word) {
        // Get Word document content
        Word.run(async (context) => {
          try {
            // Get the document title if available
            let title = "";
            try {
              const properties = context.document.properties;
              properties.load("title");
              await context.sync();
              title = properties.title;
            } catch (e) {
              console.warn("Could not get document title:", e);
            }
            
            // Get the selected text if any
            const selection = context.document.getSelection();
            selection.load("text");
            
            // Load the body with paragraphs but without list properties initially
            // This avoids the ItemNotFound error
            const body = context.document.body;
            const paragraphs = body.paragraphs;
            paragraphs.load("text");
            
            await context.sync();
            
            let content = "";
            
            // Add title if available
            if (title) {
              content += `## Document Title: ${title}\n\n`;
            }
            
            // Add selection if available
            if (selection.text && selection.text.trim().length > 0) {
              content += `## Selected Text: ${selection.text}\n`;
            }
            
            // Process paragraphs with basic formatting
            content += `## Document Content:\n\n`;
            
            // Simple approach that doesn't rely on list properties
            for (let i = 0; i < paragraphs.items.length; i++) {
              const paragraph = paragraphs.items[i];
              const text = paragraph.text.trim();
              
              if (!text) continue; // Skip empty paragraphs
              
              // Check if the text already starts with a bullet or number
              const startsWithBullet = /^[-•·○◦*]\s/.test(text);
              const startsWithNumber = /^\d+[.)]\s/.test(text);
              
              if (startsWithBullet) {
                content += `- ${text.replace(/^[-•·○◦*]\s/, '')}\n`;
              } else if (startsWithNumber) {
                content += `${text}\n`;
              } else {
                content += `${text}\n`;
              }
            }
            
            resolve(content);
          } catch (error) {
            console.error("Error in Word.run:", error);
            
            // Fallback approach if the main approach fails
            try {
              // Get just the document text without trying to preserve formatting
              const body = context.document.body;
              body.load("text");
              const selection = context.document.getSelection();
              selection.load("text");
              
              await context.sync();
              
              let fallbackContent = "";
              
              // Add selection if available
              if (selection.text && selection.text.trim().length > 0) {
                fallbackContent += `## Selected Text: ${selection.text}\n`;
              }
              
              // Add body text
              fallbackContent += `## Document Content:\n${body.text}\n\n`;
              
              resolve(fallbackContent);
            } catch (fallbackError) {
              console.error("Fallback approach also failed:", fallbackError);
              resolve("Error retrieving document content. Please try again or select specific text.");
            }
          }
        }).catch(error => {
          console.error("Error getting Word content:", error);
          resolve("");
        });
      } else if (officeAppType === Office.HostType.PowerPoint) {
        // For PowerPoint, we'll combine both approaches:
        // 1. Get the selected text using Office.js API
        // 2. Get the full slide content using OOXML
        
        // First, get the selected text
        let selectedText = "";
        let presentationTitle = "";
        let slideIndex = "";
        
        // Get selected text and basic slide info
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, 
          function(selectionResult) {
            if (selectionResult.status === Office.AsyncResultStatus.Succeeded && selectionResult.value) {
              selectedText = selectionResult.value;
            }
            
            // Get presentation title
            Office.context.document.getFilePropertiesAsync(function(propsResult) {
              if (propsResult.status === Office.AsyncResultStatus.Succeeded && propsResult.value.title) {
                presentationTitle = propsResult.value.title;
              }
              
              // Get current slide number
              Office.context.document.getActiveViewAsync(function(activeViewResult) {
                if (activeViewResult.status === Office.AsyncResultStatus.Succeeded) {
                  slideIndex = activeViewResult.value.slideIndex;
                }
                
                // Now get the full slide content using OOXML
                getPowerPointContentFromOOXML()
                  .then(ooxmlContent => {
                    let finalContent = "";
                    
                    // Add presentation title if available
                    if (presentationTitle) {
                      finalContent += `## Presentation Title: ${presentationTitle}\n`;
                    }
                    
                    // Add current slide info if available
                    if (slideIndex) {
                      finalContent += `## Current Slide: ${slideIndex}\n\n`;
                    }
                    
                    // Add selected text if available
                    if (selectedText) {
                      finalContent += `## Selected Text: ${selectedText}\n`;
                    }
                    
                    // Add OOXML content, but remove duplicate presentation title and slide number
                    // since we've already added them
                    if (ooxmlContent) {
                      let contentLines = ooxmlContent.split('\n');
                      let filteredLines = contentLines.filter(line => 
                        !line.startsWith('Presentation Title:') && 
                        !line.startsWith('Current Slide:')
                      );
                      
                      finalContent += filteredLines.join('\n');
                    }
                    
                    // If OOXML approach failed or returned empty content, fall back to Office.js
                    if (!ooxmlContent || !ooxmlContent.includes('## Slide Content')) {
                      console.log("OOXML approach didn't provide slide content, falling back to Office.js");
                      getSlideContentUsingOfficeJS(function(officeJsContent) {
                        // Add Office.js content, but avoid duplicating information
                        if (officeJsContent) {
                          let contentLines = officeJsContent.split('\n');
                          let filteredLines = contentLines.filter(line => 
                            !line.startsWith('Presentation Title:') && 
                            !line.startsWith('Current Slide:') &&
                            !line.startsWith('Selected Text:')
                          );
                          
                          finalContent += filteredLines.join('\n');
                        }
                        
                        resolve(finalContent);
                      });
                    } else {
                      resolve(finalContent);
                    }
                  })
                  .catch(error => {
                    console.error("Error with OOXML approach:", error);
                    // Fall back to the original method
                    getPowerPointContentUsingOfficeJS(resolve);
                  });
              });
            });
          }
        );
      } else {
        resolve("");
      }
    } catch (error) {
      console.error("Error accessing document content:", error);
      resolve("");
    }
  });
}

/**
 * Gets PowerPoint content using the Office.js API (original method)
 * @param {Function} resolve - The resolve function from the parent Promise
 */
function getPowerPointContentUsingOfficeJS(resolve) {
  try {
    // First try to get the selected text
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, 
      function(selectionResult) {
        const hasSelectedText = selectionResult.status === Office.AsyncResultStatus.Succeeded && selectionResult.value;
        
        // Get information about the presentation
        Office.context.document.getFilePropertiesAsync(function(propsResult) {
          let presentationInfo = "";
          if (propsResult.status === Office.AsyncResultStatus.Succeeded) {
            if (propsResult.value.title) {
              presentationInfo += `Presentation Title: ${propsResult.value.title}\n\n`;
            }
          }
          
          // Get the current slide
          Office.context.document.getActiveViewAsync(function(activeViewResult) {
            if (activeViewResult.status === Office.AsyncResultStatus.Succeeded) {
              const slideIndex = activeViewResult.value.slideIndex;
              
              // Try to get all content from the slide
              Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, 
                function(slideRangeResult) {
                  let slideContent = "";
                  
                  // Add presentation info
                  slideContent += presentationInfo;
                  
                  // Add selected text if available
                  if (hasSelectedText) {
                    slideContent += `Selected Text: ${selectionResult.value}\n\n`;
                  }
                  
                  // Add slide information
                  slideContent += `Current Slide: ${slideIndex}\n\n`;
                  
                  // Try to get all text from the current slide using a different approach
                  Office.context.document.getSelectedDataAsync(Office.CoercionType.Html, 
                    function(htmlResult) {
                      if (htmlResult.status === Office.AsyncResultStatus.Succeeded && htmlResult.value) {
                        // Extract text from HTML
                        const tempDiv = document.createElement('div');
                        tempDiv.innerHTML = htmlResult.value;
                        const extractedText = tempDiv.textContent || tempDiv.innerText || "";
                        
                        if (extractedText.trim()) {
                          slideContent += `## Slide Content: ${extractedText.trim()}\n\n`;
                        }
                      }
                      
                      // Try to get all shapes on the slide
                      try {
                        // Use Office.js API to get all shapes on the current slide
                        PowerPoint.run(async function(context) {
                          const slide = context.presentation.getActiveSlide();
                          const shapes = slide.shapes;
                          shapes.load("items");
                          
                          await context.sync();
                          
                          let shapeTexts = [];
                          for (let i = 0; i < shapes.items.length; i++) {
                            if (shapes.items[i].hasTextFrame) {
                              const textFrame = shapes.items[i].textFrame;
                              textFrame.load("textRange");
                              await context.sync();
                              
                              const textRange = textFrame.textRange;
                              textRange.load("text");
                              await context.sync();
                              
                              if (textRange.text) {
                                shapeTexts.push(textRange.text);
                              }
                            }
                          }
                          
                          if (shapeTexts.length > 0) {
                            slideContent += `Slide Text Elements:\n${shapeTexts.join("\n")}\n\n`;
                          }
                          
                          resolve(slideContent);
                        }).catch(function(error) {
                          // PowerPoint API might not be available in all versions
                          console.log("PowerPoint API not available, falling back to basic content");
                          
                          // If we can't get shape text, just return what we have so far
                          if (!slideContent.includes('## Slide Content: ') && !hasSelectedText) {
                            slideContent += "Note: For best results, please select the text you want to include from the slide.\n";
                          }
                          
                          resolve(slideContent);
                        });
                      } catch (e) {
                        console.error("Error getting PowerPoint shapes:", e);
                        
                        // If we can't get shape text, just return what we have so far
                        if (!slideContent.includes('## Slide Content: ') && !hasSelectedText) {
                          slideContent += "Note: For best results, please select the text you want to include from the slide.\n";
                        }
                        
                        resolve(slideContent);
                      }
                    }
                  );
                }
              );
            } else {
              // Fallback if we can't get the active view
              let content = presentationInfo;
              
              // Add selected text if available
              if (hasSelectedText) {
                content += `Selected Text: ${selectionResult.value}\n\n`;
              } else {
                content += "PowerPoint content: Unable to access slide content. Please select text to include specific content.";
              }
              
              resolve(content);
            }
          });
        });
      }
    );
  } catch (error) {
    console.error("Error accessing PowerPoint content:", error);
    resolve("Error accessing PowerPoint content. Please select text to include specific content.");
  }
}

/**
 * Gets slide content using Office.js API (helper function)
 * @param {Function} callback - Callback function to receive the content
 */
function getSlideContentUsingOfficeJS(callback) {
  try {
    // Try to get all text from the current slide using HTML approach
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Html, 
      function(htmlResult) {
        let slideContent = "";
        
        if (htmlResult.status === Office.AsyncResultStatus.Succeeded && htmlResult.value) {
          // Extract text from HTML
          const tempDiv = document.createElement('div');
          tempDiv.innerHTML = htmlResult.value;
          const extractedText = tempDiv.textContent || tempDiv.innerText || "";
          
          if (extractedText.trim()) {
            slideContent += `## Slide Content: ${extractedText.trim()}\n\n`;
          }
        }
        
        // Try to get all shapes on the slide
        try {
          // Use Office.js API to get all shapes on the current slide
          PowerPoint.run(async function(context) {
            const slide = context.presentation.getActiveSlide();
            const shapes = slide.shapes;
            shapes.load("items");
            
            await context.sync();
            
            let shapeTexts = [];
            for (let i = 0; i < shapes.items.length; i++) {
              if (shapes.items[i].hasTextFrame) {
                const textFrame = shapes.items[i].textFrame;
                textFrame.load("textRange");
                await context.sync();
                
                const textRange = textFrame.textRange;
                textRange.load("text");
                await context.sync();
                
                if (textRange.text) {
                  shapeTexts.push(textRange.text);
                }
              }
            }
            
            if (shapeTexts.length > 0) {
              slideContent += `Slide Text Elements:\n${shapeTexts.join("\n")}\n\n`;
            }
            
            callback(slideContent);
          }).catch(function(error) {
            // PowerPoint API might not be available in all versions
            console.log("PowerPoint API not available, returning basic content");
            callback(slideContent);
          });
        } catch (e) {
          console.error("Error getting PowerPoint shapes:", e);
          callback(slideContent);
        }
      }
    );
  } catch (error) {
    console.error("Error in getSlideContentUsingOfficeJS:", error);
    callback("");
  }
}

/**
 * Gets the active slide number in PowerPoint
 * @returns {Promise<number>} - The active slide number or null if not available
 */
function getActiveSlideNumber() {
  return new Promise((resolve) => {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function(result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        if (result.value && result.value.slides && result.value.slides.length > 0) {
          console.log("Active slide number: " + result.value.slides[0].index);
          resolve(result.value.slides[0].index);
        } else {
          console.log("No slide selected or unable to determine slide number");
          resolve(null);
        }
      } else {
        console.error("Error getting slide number: " + (result.error ? result.error.message : "Unknown error"));
        resolve(null);
      }
    });
  });
}

/**
 * Gets PowerPoint content using the OOXML approach
 * @returns {Promise<string>} - The extracted slide content
 */
async function getPowerPointContentFromOOXML() {
  return new Promise(async (resolve, reject) => {
    try {
      // Ensure dependencies are loaded
      try {
        await Promise.all([loadJSZip(), loadXMLDOM()]);
      } catch (error) {
        console.error("Failed to load dependencies:", error);
        resolve(""); // Return empty string to fall back to Office.js method
        return;
      }
      
      const docType = Office.FileType.Compressed;
      const docParams = { sliceSize: 65536 };
      
      Office.context.document.getFileAsync(docType, docParams, async (result) => {
        if (result.status === "succeeded") {
          const file = result.value;
          const sliceCount = file.sliceCount;
          
          let processedSlices = 0;
          const sliceData = [];
          
          // Get the active slide number using the new function
          const currentSlide = await getActiveSlideNumber();
          
          // Function to process each slice
          function processNextSlice(index) {
            file.getSliceAsync(index, (sliceResult) => {
              if (sliceResult.status === "succeeded") {
                sliceData[sliceResult.value.index] = sliceResult.value.data;
                processedSlices++;
                
                if (processedSlices === sliceCount) {
                  // All slices processed, combine and parse
                  file.closeAsync();
                  processOOXMLData(sliceData, currentSlide)
                    .then(content => {
                      resolve(content);
                    })
                    .catch(error => {
                      console.error("Error processing OOXML:", error);
                      resolve("");
                    });
                } else {
                  // Process next slice
                  processNextSlice(index + 1);
                }
              } else {
                console.error(`Error getting slice ${index}:`, sliceResult.error);
                file.closeAsync();
                resolve("");
              }
            });
          }
          
          // Start processing from the first slice
          if (sliceCount > 0) {
            processNextSlice(0);
          } else {
            console.warn("No slices found in the file");
            file.closeAsync();
            resolve("");
          }
        } else {
          console.error("Failed to get file:", result.error);
          resolve("");
        }
      });
    } catch (error) {
      console.error("Error in getPowerPointContentFromOOXML:", error);
      resolve("");
    }
  });
}

/**
 * Process the OOXML data to extract slide content
 * @param {Array} sliceData - Array of data slices from the file
 * @param {number} currentSlide - The current slide number
 * @returns {Promise<string>} - The extracted slide content
 */
async function processOOXMLData(sliceData, currentSlide) {
  try {
    // Combine all slices into a single array
    let combinedData = [];
    for (let i = 0; i < sliceData.length; i++) {
      combinedData = combinedData.concat(sliceData[i]);
    }
    
    // Convert to string for processing
    let dataString = "";
    for (let j = 0; j < combinedData.length; j++) {
      dataString += String.fromCharCode(combinedData[j]);
    }
    
    // Create a Uint8Array for JSZip
    const dataArray = new Uint8Array(dataString.length);
    for (let i = 0; i < dataString.length; i++) {
      dataArray[i] = dataString.charCodeAt(i);
    }
    
    // Check if JSZip is available
    if (typeof JSZip === 'undefined') {
      console.warn("JSZip is not available. Cannot process OOXML data.");
      return "";
    }
    
    // Load the ZIP file
    const zip = new JSZip();
    const zipData = await zip.loadAsync(dataArray);
    
    // Debug: List only slide files in the ZIP (not all files)
    const slideFiles = Object.keys(zipData.files).filter(filename => 
      filename.startsWith('ppt/slides/') && 
      filename.endsWith('.xml') && 
      !filename.includes('_rels')
    );
    console.log("Available slide files:", slideFiles);
    
    // Use the provided currentSlide if available, otherwise try to get it
    let slideNumber = currentSlide;
    let slideContent = "";
    
    // If we still don't have a slide number, try the old method as fallback
    if (!slideNumber) {
      console.log("Using fallback method to get slide number");
      // Try to get the current slide number using the old method
      await new Promise(resolve => {
        Office.context.document.getActiveViewAsync(function(result) {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            slideNumber = result.value.slideIndex;
            console.log("Retrieved slide index (fallback):", slideNumber);
          } else {
            console.log("Could not get active view:", result.error);
          }
          resolve();
        });
      });
    }
    
    // Get presentation properties
    try {
      if (zipData.files["docProps/core.xml"]) {
        const corePropsXml = await zipData.file("docProps/core.xml").async("text");
        
        // Use the appropriate DOMParser (browser's or xmldom's)
        const parser = new DOMParser();
        const coreDoc = parser.parseFromString(corePropsXml, "text/xml");
        
        // Try different tag names for the title (different OOXML versions might use different namespaces)
        let titleElements = coreDoc.getElementsByTagName("dc:title");
        if (!titleElements || titleElements.length === 0) {
          titleElements = coreDoc.getElementsByTagName("title");
        }
        
        if (titleElements && titleElements.length > 0) {
          const title = titleElements[0].textContent;
          if (title) {
            slideContent += `Presentation Title: ${title}\n\n`;
          }
        }
      }
    } catch (e) {
      console.warn("Could not extract presentation properties:", e);
    }
    
    // Add current slide info
    slideContent += `Current Slide: ${slideNumber || "Unknown"}\n\n`;
    
    // Try to get content from current slide - try multiple possible formats
    let slideFound = false;
    
    // List of possible slide filename formats to try
    const slideFormats = [];
    
    // Try both 0-based and 1-based indexing
    // Some PowerPoint implementations might use 0-based indexing internally
    const possibleIndices = [];
    
    // If slideNumber is defined and is a number, use it as the primary index
    if (slideNumber !== null && slideNumber !== undefined && !isNaN(slideNumber)) {
      possibleIndices.push(slideNumber);
      // Also try adjacent slides in case of off-by-one errors
      if (slideNumber > 1) possibleIndices.push(slideNumber - 1);
      possibleIndices.push(slideNumber + 1);
    } else {
      // If no slide number, try some defaults
      possibleIndices.push(1);
    }
    
    // Generate all possible slide paths to try
    for (const index of possibleIndices) {
      slideFormats.push(`ppt/slides/slide${index}.xml`);
    }
    
    // Also try some default names
    slideFormats.push('ppt/slides/slide1.xml');
    
    // Try each possible slide format
    for (const slideFileName of slideFormats) {
      if (zipData.files[slideFileName]) {
        console.log(`Found slide file: ${slideFileName}`);
        try {
          const slideXml = await zipData.file(slideFileName).async("text");
          
          // Use the appropriate DOMParser
          const parser = new DOMParser();
          const slideDoc = parser.parseFromString(slideXml, "text/xml");
          
          // Process the slide content with structure preservation
          const structuredContent = extractStructuredContent(slideDoc);
          
          if (structuredContent) {
            slideContent += `## Slide Content:\n${structuredContent}\n`;
            slideFound = true;
            break; // Found and processed a slide, no need to try other formats
          }
        } catch (e) {
          console.warn(`Error processing ${slideFileName}:`, e);
        }
      }
    }
    
    // If we still haven't found a slide, try a more general approach
    if (!slideFound) {
      // Look for any slide files
      const availableSlideFiles = Object.keys(zipData.files).filter(filename => 
        filename.startsWith('ppt/slides/') && filename.endsWith('.xml') && !filename.includes('_rels')
      );
      
      if (availableSlideFiles.length > 0) {
        // Try the first slide file we find
        try {
          const slideXml = await zipData.file(availableSlideFiles[0]).async("text");
          const parser = new DOMParser();
          const slideDoc = parser.parseFromString(slideXml, "text/xml");
          
          // Process the slide content with structure preservation
          const structuredContent = extractStructuredContent(slideDoc);
          
          if (structuredContent) {
            slideContent += `## Slide Content:\n`;
            slideFound = true;
          }
        } catch (e) {
          console.warn(`Error processing ${availableSlideFiles[0]}:`, e);
        }
      }
    }
    
    if (!slideFound) {
      slideContent += "Could not find slide XML data.\n";
    }
    
    return slideContent;
  } catch (error) {
    console.error("Error in processOOXMLData:", error);
    return "Error processing slide content: " + error.message;
  }
}

/**
 * Extracts structured content from a slide, preserving lists and other formatting
 * @param {Document} slideDoc - The parsed XML document of the slide
 * @returns {string} - The structured content as text
 */
function extractStructuredContent(slideDoc) {
  let content = "";
  
  try {
    // Get all paragraphs (a:p elements)
    const paragraphs = slideDoc.getElementsByTagName("a:p");
    
    if (!paragraphs || paragraphs.length === 0) {
      return null;
    }
    
    // Process each paragraph
    for (let i = 0; i < paragraphs.length; i++) {
      const paragraph = paragraphs[i];
      
      // Check if this paragraph is part of a list
      const pPr = paragraph.getElementsByTagName("a:pPr")[0];
      let isList = false;
      let isNumbered = false;
      let listLevel = 0;
      let bulletChar = "•"; // Default bullet character
      
      if (pPr) {
        // Check for list bullet element (a:buChar)
        const buChar = pPr.getElementsByTagName("a:buChar")[0];
        if (buChar) {
          isList = true;
          
          // Try to get the actual bullet character
          if (buChar.hasAttribute("char")) {
            bulletChar = buChar.getAttribute("char");
          }
          
          // Check for list level (lvl attribute)
          if (pPr.hasAttribute("lvl")) {
            listLevel = parseInt(pPr.getAttribute("lvl"), 10);
          }
        }
        
        // Also check for numbered list (a:buAutoNum)
        const buAutoNum = pPr.getElementsByTagName("a:buAutoNum")[0];
        if (buAutoNum) {
          isList = true;
          isNumbered = true;
          
          // Check for list level (lvl attribute)
          if (pPr.hasAttribute("lvl")) {
            listLevel = parseInt(pPr.getAttribute("lvl"), 10);
          }
          
          // Try to get the numbering type
          if (buAutoNum.hasAttribute("type")) {
            // Different types: arabicPeriod, romanLcPeriod, etc.
            // We'll just use "1." for simplicity
          }
        }
        
        // Also check for bullet none (a:buNone) which explicitly disables bullets
        const buNone = pPr.getElementsByTagName("a:buNone")[0];
        if (buNone) {
          isList = false;
        }
      }
      
      // Get all text runs in this paragraph
      const textRuns = paragraph.getElementsByTagName("a:t");
      let paragraphText = "";
      
      // Combine all text runs in this paragraph
      for (let j = 0; j < textRuns.length; j++) {
        const text = textRuns[j].textContent;
        if (text && text.trim()) {
          paragraphText += text + " ";
        }
      }
      
      paragraphText = paragraphText.trim();
      
      // Add the paragraph with appropriate formatting
      if (paragraphText) {
        // Check if the text already starts with a bullet or number
        // This handles cases where the bullet is part of the text content
        const startsWithBullet = /^[-•·○◦*]\s/.test(paragraphText);
        const startsWithNumber = /^\d+[.)]\s/.test(paragraphText);
        
        if (isList && !startsWithBullet && !startsWithNumber) {
          // Add indentation based on list level
          const indent = "  ".repeat(listLevel);
          
          if (isNumbered) {
            // For numbered lists
            content += `${indent}1. ${paragraphText}\n`;
          } else {
            // For bullet lists - use the actual bullet character if possible
            // Convert to a standard markdown bullet for consistency
            content += `${indent}- ${paragraphText}\n`;
          }
        } else if (startsWithBullet) {
          // Text already has a bullet, just add indentation
          const indent = "  ".repeat(listLevel);
          // Ensure it uses a standard markdown bullet
          content += `${indent}- ${paragraphText.replace(/^[-•·○◦*]\s/, '')}\n`;
        } else if (startsWithNumber) {
          // Text already has a number, just add indentation
          const indent = "  ".repeat(listLevel);
          content += `${indent}${paragraphText}\n`;
        } else {
          content += `${paragraphText}\n`;
        }
      }
    }
    
    return content;
  } catch (error) {
    console.error("Error extracting structured content:", error);
    return null;
  }
}

/**
 * Adds a user message to the chat interface
 * @param {string} message - The message text
 */
function addUserMessage(message) {
  const messageContainer = document.createElement('div');
  messageContainer.className = 'message-container user-container';
  
  const messageHeader = document.createElement('div');
  messageHeader.className = 'message-header user-header';
  messageHeader.textContent = 'You';
  
  const messageElement = document.createElement('div');
  messageElement.className = 'message user-message';
  
  // Create a div for the content with markdown
  const messageContent = document.createElement('div');
  messageContent.className = 'message-content';
  messageContent.innerHTML = marked.parse(escapeHtml(message));
  
  messageElement.appendChild(messageContent);
  messageContainer.appendChild(messageHeader);
  messageContainer.appendChild(messageElement);
  chatMessages.appendChild(messageContainer);
  
  // Add language badges to code blocks
  addLanguageBadges(messageElement);
  
  // Scroll to bottom
  chatMessages.scrollTop = chatMessages.scrollHeight;
}

/**
 * Adds a bot message to the chat interface
 * @param {string} message - The message text
 */
function addBotMessage(message) {
  const messageContainer = document.createElement('div');
  messageContainer.className = 'message-container';
  
  const messageHeader = document.createElement('div');
  messageHeader.className = 'message-header';
  messageHeader.textContent = 'Cria';
  
  const messageElement = document.createElement('div');
  messageElement.className = 'message bot-message';
  
  // Create a div for the content with markdown
  const messageContent = document.createElement('div');
  messageContent.className = 'message-content';
  
  try {
    messageContent.innerHTML = marked.parse(message);
  } catch (e) {
    console.error("Error parsing markdown:", e);
    messageContent.textContent = message;
  }
  
  messageElement.appendChild(messageContent);
  messageContainer.appendChild(messageHeader);
  messageContainer.appendChild(messageElement);
  chatMessages.appendChild(messageContainer);
  
  // Add language badges to code blocks
  addLanguageBadges(messageElement);
  
  // Scroll to bottom
  chatMessages.scrollTop = chatMessages.scrollHeight;
  
  // Initialize Prism highlighting on the new content
  Prism.highlightAllUnder(messageElement);
}

// ===== CODE BLOCK UTILITIES =====

/**
 * Adds language badges and copy buttons to code blocks
 * @param {HTMLElement} element - The element containing code blocks
 */
function addLanguageBadges(element) {
  element.querySelectorAll('pre code').forEach((block) => {
    const language = block.className.match(/language-(\w+)/)?.[1] || 'javascript';
    const displayName = getLanguageDisplayName(language);
    
    // Create header element
    const header = document.createElement('div');
    header.className = 'code-header';
    
    // Add language name
    const langName = document.createElement('span');
    langName.textContent = displayName;
    header.appendChild(langName);
    
    // Add copy button
    const copyButton = document.createElement('button');
    copyButton.className = 'copy-button';
    copyButton.innerHTML = '<i class="bi bi-clipboard"></i> Copy';
    copyButton.addEventListener('click', function() {
      copyCodeToClipboard(block);
    });
    header.appendChild(copyButton);
    
    // Insert header before the pre element
    block.parentElement.parentElement.insertBefore(header, block.parentElement);
  });
}

/**
 * Copies code block content to clipboard
 * @param {HTMLElement} codeBlock - The code block element
 */
function copyCodeToClipboard(codeBlock) {
  const code = codeBlock.textContent;
  navigator.clipboard.writeText(code).then(() => {
    showCopyFeedback(codeBlock, 'Copied!');
    
    // Change the button text temporarily
    const copyButton = codeBlock.parentElement.previousSibling.querySelector('.copy-button');
    const originalHTML = copyButton.innerHTML;
    copyButton.innerHTML = '<i class="bi bi-check-lg"></i> Copied';
    
    setTimeout(() => {
      copyButton.innerHTML = originalHTML;
    }, 1500);
    
  }).catch(err => {
    showCopyFeedback(codeBlock, 'Failed to copy');
    console.error('Could not copy text: ', err);
  });
}

/**
 * Shows feedback when code is copied
 * @param {HTMLElement} element - The element that was copied
 * @param {string} message - The feedback message
 */
function showCopyFeedback(element, message) {
  // Create feedback element if it doesn't exist
  let feedback = document.querySelector('.copy-feedback');
  if (!feedback) {
    feedback = document.createElement('div');
    feedback.className = 'copy-feedback';
    document.body.appendChild(feedback);
  }
  
  // Position the feedback near the element
  const rect = element.getBoundingClientRect();
  feedback.style.top = `${rect.top - 30}px`;
  feedback.style.left = `${rect.left + rect.width / 2 - 40}px`;
  feedback.textContent = message;
  
  // Show the feedback
  feedback.classList.add('show-feedback');
  
  // Hide after a delay
  setTimeout(() => {
    feedback.classList.remove('show-feedback');
  }, 1500);
}

// ===== HELPER FUNCTIONS =====

/**
 * Escapes HTML in user input while preserving markdown code blocks
 * @param {string} text - The text to escape
 * @returns {string} - HTML-escaped text with preserved markdown
 */
function escapeHtml(text) {
  // We want to escape HTML but preserve markdown code blocks
  const codeBlocks = [];
  // Replace triple backtick code blocks with placeholders
  text = text.replace(/```([\s\S]*?)```/g, function(match) {
    const placeholder = `__CODE_BLOCK_${codeBlocks.length}__`;
    codeBlocks.push(match);
    return placeholder;
  });
  
  // Replace inline code with placeholders
  const inlineCode = [];
  text = text.replace(/`([^`]+)`/g, function(match) {
    const placeholder = `__INLINE_CODE_${inlineCode.length}__`;
    inlineCode.push(match);
    return placeholder;
  });
  
  // Escape HTML
  const escapeMap = {
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#039;'
  };
  text = text.replace(/[&<>"']/g, function(m) { return escapeMap[m]; });
  
  // Restore code blocks
  codeBlocks.forEach((block, i) => {
    text = text.replace(`__CODE_BLOCK_${i}__`, block);
  });
  
  // Restore inline code
  inlineCode.forEach((code, i) => {
    text = text.replace(`__INLINE_CODE_${i}__`, code);
  });
  
  return text;
}

/**
 * Formats conversation history for the API
 * @param {Array} history - The conversation history
 * @param {string} documentContent - Optional document content to include
 * @returns {string} - Formatted prompt for the API
 */
function formatConversationForAPI(history, documentContent = "") {
  // Generate the system prompt
  const systemPrompt = generateSystemPrompt();
  
  // Start with the system prompt
  let formattedPrompt = `System: ${systemPrompt}\n\n`;
  
  // Add conversation history
  history.forEach(message => {
    const role = message.role === "user" ? "User" : "Assistant";
    formattedPrompt += `${role}: ${message.content}\n`;
  });
  
  // Add document content if available
  if (documentContent) {
    // Toggle prefix based on Office app type
    let prefix = "# Attached document:\n";
    if (officeAppType === Office.HostType.PowerPoint) {
      prefix = "# Attached slide:\n";
    }
    formattedPrompt += `\n${prefix}${documentContent}\n\n`;
  }
  console.log(formattedPrompt);
  return formattedPrompt;
}

/**
 * Lists all slide files in the PowerPoint presentation for debugging purposes
 * This can help identify the correct file structure
 */
function debugListPresentationFiles() {
  console.log("Listing slide files in the presentation...");
  
  try {
    // Ensure dependencies are loaded
    Promise.all([loadJSZip(), loadXMLDOM()])
      .then(() => {
        const docType = Office.FileType.Compressed;
        const docParams = { sliceSize: 65536 };
        
        Office.context.document.getFileAsync(docType, docParams, (result) => {
          if (result.status === "succeeded") {
            const file = result.value;
            const sliceCount = file.sliceCount;
            console.log(`File retrieved. Processing ${sliceCount} slices...`);
            
            let processedSlices = 0;
            const sliceData = [];
            
            // Function to process each slice
            function processNextSlice(index) {
              file.getSliceAsync(index, (sliceResult) => {
                if (sliceResult.status === "succeeded") {
                  sliceData[sliceResult.value.index] = sliceResult.value.data;
                  processedSlices++;
                  
                  if (processedSlices === sliceCount) {
                    // All slices processed, combine and parse
                    file.closeAsync();
                    
                    // Combine all slices into a single array
                    let combinedData = [];
                    for (let i = 0; i < sliceData.length; i++) {
                      combinedData = combinedData.concat(sliceData[i]);
                    }
                    
                    // Convert to string for processing
                    let dataString = "";
                    for (let j = 0; j < combinedData.length; j++) {
                      dataString += String.fromCharCode(combinedData[j]);
                    }
                    
                    // Create a Uint8Array for JSZip
                    const dataArray = new Uint8Array(dataString.length);
                    for (let i = 0; i < dataString.length; i++) {
                      dataArray[i] = dataString.charCodeAt(i);
                    }
                    
                    // Load the ZIP file
                    const zip = new JSZip();
                    zip.loadAsync(dataArray)
                      .then(zipData => {
                        // List only slide files
                        const slideFiles = Object.keys(zipData.files).filter(name => 
                          name.startsWith('ppt/slides/') && 
                          name.endsWith('.xml') && 
                          !name.includes('_rels')
                        );
                        console.log("Slide files:", slideFiles);
                        
                        // Try to get the current slide number
                        Office.context.document.getActiveViewAsync(function(viewResult) {
                          if (viewResult.status === Office.AsyncResultStatus.Succeeded) {
                            console.log("Current slide index:", viewResult.value.slideIndex);
                            
                            // Try to get content from the current slide
                            const currentSlide = viewResult.value.slideIndex;
                            const currentSlideFile = `ppt/slides/slide${currentSlide}.xml`;
                            
                            if (zipData.files[currentSlideFile]) {
                              console.log(`Current slide file exists: ${currentSlideFile}`);
                              
                              zipData.file(currentSlideFile).async("text")
                                .then(slideXml => {
                                  const parser = new DOMParser();
                                  const slideDoc = parser.parseFromString(slideXml, "text/xml");
                                  const textElements = slideDoc.getElementsByTagName("a:t");
                                  
                                  if (textElements && textElements.length > 0) {
                                    console.log(`Found ${textElements.length} text elements in current slide`);
                                    
                                    // Show first few text elements as sample
                                    const sampleCount = Math.min(5, textElements.length);
                                    console.log(`Sample text content (first ${sampleCount} elements):`);
                                    
                                    for (let i = 0; i < sampleCount; i++) {
                                      const text = textElements[i].textContent;
                                      if (text && text.trim()) {
                                        console.log(`- "${text}"`);
                                      }
                                    }
                                  } else {
                                    console.log("No text elements found in current slide");
                                  }
                                })
                                .catch(error => {
                                  console.error("Error reading current slide:", error);
                                });
                            } else {
                              console.log(`Current slide file not found: ${currentSlideFile}`);
                            }
                          } else {
                            console.log("Could not get current slide index");
                          }
                        });
                      })
                      .catch(error => {
                        console.error("Error processing ZIP:", error);
                      });
                  } else {
                    // Process next slice
                    processNextSlice(index + 1);
                  }
                } else {
                  file.closeAsync();
                  console.error("Error getting slice:", sliceResult.error);
                }
              });
            }
            
            // Start processing from the first slice
            if (sliceCount > 0) {
              processNextSlice(0);
            } else {
              console.warn("No slices found in the file");
              file.closeAsync();
            }
          } else {
            console.error("Failed to get file:", result.error);
          }
        });
      })
      .catch(error => {
        console.error("Failed to load dependencies:", error);
      });
  } catch (error) {
    console.error("Error in debugListPresentationFiles:", error);
  }
}