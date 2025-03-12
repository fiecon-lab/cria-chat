/**
 * Cria Chat - A markdown-enabled chat interface
 * This script handles the chat functionality, markdown parsing, and API integration
 */

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

// DOM elements
const chatMessages = document.getElementById('chat-messages');
const messageInput = document.getElementById('message-input');
const sendButton = document.getElementById('send-button');
const typingIndicator = document.getElementById('typing-indicator');
const newChatButton = document.getElementById('new-chat-button');

// API details
const API_URL = "https://cria-api.fiecon.com/api/generate";
const API_KEY = "0a2e6ef6-4a96-406f-888e-865a8c5a7209";
const API_MODEL = "llama3.2-vision:latest";

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

// ===== EVENT LISTENERS =====

// Initialize with a welcome message
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
};

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
newChatButton.addEventListener('click', startNewConversation);

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
    // Prepare the prompt with conversation history
    const prompt = formatConversationForAPI(conversationHistory);
    
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
 * @returns {string} - Formatted prompt for the API
 */
function formatConversationForAPI(history) {
  // For simple API that doesn't support chat format natively,
  // we'll format the conversation as a text prompt
  let formattedPrompt = "";
  
  history.forEach(message => {
    const role = message.role === "user" ? "User" : "Assistant";
    formattedPrompt += `${role}: ${message.content}\n`;
  });
  
  // Add the final prompt for the assistant to respond
  formattedPrompt += "Assistant:";
  
  return formattedPrompt;
}