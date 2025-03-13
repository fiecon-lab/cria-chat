# Cria Chat

## Overview
Cria Chat is a Microsoft Office add-in that provides an AI-powered chat interface within Word, Excel, and PowerPoint. Developed by FIECON, a consultancy specializing in health economics and market access, Cria Chat enables users to interact with an AI assistant directly within their Office applications.

## Features

### Core Functionality
- **AI-Powered Assistance**: Integrates with FIECON's AI model to provide intelligent responses to user queries
- **Markdown Support**: Renders markdown formatting in chat messages, including code blocks with syntax highlighting
- **Document Context Awareness**: Can access and analyze the content of the active document when permitted
- **Office Integration**: Works seamlessly within Microsoft Word, Excel, and PowerPoint

### Office-Specific Features
- **Word Integration**: 
  - Access document content including selected text
  - Analyze document structure and formatting
  
- **PowerPoint Integration**: 
  - Access slide content and structure
  - Extract text from the active slide
  - Support for selected text on slides
  
- **Excel Integration**: 
  - Basic support (document content access is limited)

### Technical Features
- **Code Highlighting**: Uses Prism.js to provide syntax highlighting for code blocks
- **Dynamic Content Loading**: Loads dependencies like JSZip and XML parsing libraries as needed
- **Responsive Design**: Adapts to different screen sizes and Office environments
- **OOXML Processing**: Advanced capability to extract content from Office documents using Open Office XML format

## How It Works

1. **Initialization**: The add-in initializes when loaded in an Office application, detecting the host application type (Word, Excel, or PowerPoint)
2. **User Interface**: Provides a chat interface where users can type messages and receive responses
3. **Document Access**: When enabled, the add-in can access document content to provide context-aware responses
4. **API Integration**: Communicates with FIECON's API to generate responses based on the conversation history and document context
5. **Rendering**: Processes and renders responses with markdown formatting and code syntax highlighting

## Technical Implementation

- **Frontend**: HTML, CSS, JavaScript with Bootstrap for responsive design
- **Office Integration**: Office.js API for interacting with Office applications
- **Markdown Processing**: Uses marked.js for parsing markdown
- **Syntax Highlighting**: Prism.js for code block highlighting
- **Document Processing**: 
  - Office.js API for basic document access
  - JSZip and XML parsing for advanced OOXML processing in PowerPoint

## Security and Privacy

- Document content is only accessed when explicitly permitted by the user
- The add-in communicates with FIECON's secure API endpoint
- No document content is stored permanently

## Requirements

- Microsoft Office (Word, Excel, or PowerPoint)
- Internet connection for API communication

## Development

This project is maintained by FIECON Labs. The add-in is built using standard web technologies and the Office Add-in framework.

---

© FIECON Ltd. All rights reserved.