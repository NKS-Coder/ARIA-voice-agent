ARIA — Voice AI Agent

ARIA is a browser-based voice AI assistant that enables real-time interaction with large language models using natural speech.

The system integrates Groq’s Llama models, ElevenLabs voice synthesis, and the Web Speech API to deliver a responsive conversational AI experience.

This project demonstrates how modern voice assistants can be built using lightweight web technologies and external AI APIs.



The assistant is deployed using GitHub Pages.

https://NKS-Coder.github.io/ARIA-voice-agent/index.html


Open the link and interact with the assistant directly in your browser.

Interface Preview



![ARIA Interface](screenshot.png)

The interface allows users to activate the assistant using microphone input and receive synthesized voice responses from the AI model.

Features

Real-time voice interaction using the Web Speech API

Fast AI inference powered by Groq Llama models

Human-like voice responses using ElevenLabs

Multiple assistant personas for different use cases

Automation-ready architecture compatible with n8n workflows

Fully browser-based implementation

Assistant Personas
Mode	Purpose
General	Everyday tasks and questions
Sales	Lead generation and conversational selling
Support	Troubleshooting and technical assistance
Research	Information retrieval and analysis
System Architecture
User Voice
   ↓
Web Speech API
   ↓
Groq LLM (Llama 3.3)
   ↓
AI Response
   ↓
ElevenLabs Voice Synthesis
   ↓
Audio Playback

This architecture enables low-latency conversational AI directly within the browser.

Technology Stack
Layer	Technology
AI Model	Groq (Llama 3.3)
Speech Recognition	Web Speech API
Voice Synthesis	ElevenLabs
Frontend	HTML / JavaScript
Hosting	GitHub Pages
Automation Integration	n8n
