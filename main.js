/* -------------------------------------------------------------------------- */
/*                                 Constants                                  */
/* -------------------------------------------------------------------------- */
const MARKER_A = '🅰️'; // Cursor Selection Start
const MARKER_B = '🅱️'; // Cursor Selection End
const MAX_HISTORY = 30;
const CURSOR_SYMBOL = '◉';

const DEFAULT_SYSTEM_PROMPT = `You are a helpful text editing assistant.
The user will provide text containing two cursor markers:
🅰️ (start of selection/cursor) and 🅱️ (end of selection/cursor). 
If 🅰️ and 🅱️ are adjacent, it represents a caret position. If they surround text, it represents a selection.
Your task is to follow the user's instruction to modify the text.
Return ONLY the fully updated text content. Do not include any explanations or markdown formatting unless requested.
By default, any new text should be inserted at a cursor position, replacing the selected text if any is selected.
Your output should include the cursor position and selection, i.e. the special characters 🅰️ and 🅱️. In case you just inserted text,
by default the cursor should be at the end of the newly inserted text with nothing selected.

USER_INSRUCTION:
`;

const DEFAULT_CONFIG = `1. Commands (start with #)
process=#process
execute=#execute
undo=#undo
redo=#redo
stop=#stop
discard=#discard

2. Substitutions
period|full stop=.
colon=:
semicolon=;
exclamation mark=!
question mark=?
comma=,
new line|enter|new paragraph=\\n
(smile|smiling|smiley) emoji=🙂
heart emoji=❤️
laughing emoji=😂
crying emoji=😭
(like|thumbs up) emoji=👍
dislike emoji=👎
angry emoji=😠
sad emoji=😢
happy emoji=😊
url=TODO_ADD_LINK_HERE_LATER
(open|left) (parenthesis|parents)=(
(close|right) (parenthesis|parents)=)
double quote="
single quote='

3. Regex Operations (trigger=match_regex:::replacement)
# Use 🅰️ for Cursor Start and 🅱️ for Cursor End
space=🅰️[\\s\\S]*?🅱️::: 🅰️🅱️
# Deletes the word immediately before the cursor/selection
delete|backspace=(\\S+\\s*)?🅰️[\\s\\S]*?🅱️:::🅰️🅱️
# Deletes the sentence segment immediately before the cursor
sentence delete=[^.!?]+[.!?]*\\s*🅰️[\\s\\S]*?🅱️:::🅰️🅱️
# Deletes selection
selection delete=🅰️[\\s\\S]*?🅱️:::🅰️🅱️
# Clears the entire document
clear all=[\\s\\S]*:::🅰️🅱️
# Clear spaces before cursor
clear space=[ \\t]*🅰️([\\s\\S]*?)🅱️[ \\t]*:::🅰️$1🅱️

# Move Up (to start of previous line)
move up=(^|[\\s\\S]*\\n)([^\\n]*)\\n([^\\n]*)🅰️([\\s\\S]*?)🅱️([^\\n]*)([\\s\\S]*):::$1🅰️🅱️$2\\n$3$4$5$6
# Move Down (to start of next line)
move down=(^|[\\s\\S]*\\n)([^\\n]*)🅰️([\\s\\S]*?)🅱️([^\\n]*)\\n([^\\n]*)([\\s\\S]*):::$1$2$3$4\\n🅰️🅱️$5$6
# Move to Start of Line
move to start( of line)?=(^|[\\s\\S]*\\n)([^\\n]*)🅰️([\\s\\S]*?)🅱️([\\s\\S]*):::$1🅰️🅱️$2$3$4
# Move to End of Line
move to end( of line)?=(^|[\\s\\S]*\\n)([^\\n]*)🅰️([\\s\\S]*?)🅱️([^\\n]*)([\\s\\S]*):::$1$2$3$4🅰️🅱️$5
# Move to Top (Start of Text)
move to top=^([\\s\\S]*)🅰️([\\s\\S]*?)🅱️([\\s\\S]*)$:::🅰️🅱️$1$2$3
# Move to Bottom (End of Text)
move to bottom=^([\\s\\S]*)🅰️([\\s\\S]*?)🅱️([\\s\\S]*)$:::$1$2$3🅰️🅱️
`;

/* -------------------------------------------------------------------------- */
/*                                DOM Elements                                */
/* -------------------------------------------------------------------------- */
const toggleButton = document.getElementById('toggleButton');
const processButton = document.getElementById('processButton');
const executeButton = document.getElementById('executeButton');
const discardButton = document.getElementById('discardButton');
const undoButton = document.getElementById('undoButton');
const redoButton = document.getElementById('redoButton');
const textBox = document.getElementById('textBox');
const pendingTextSpan = document.getElementById('pendingText');
const statusDiv = document.getElementById('status');
const sidebar = document.getElementById('sidebar');
const sidebarToggle = document.getElementById('sidebarToggle');
const configEditor = document.getElementById('configEditor');
const saveConfigButton = document.getElementById('saveConfigButton');
const resetConfigButton = document.getElementById('resetConfigButton');
const groqApiKeyInput = document.getElementById('groqApiKeyInput');
const geminiApiKeyInput = document.getElementById('geminiApiKeyInput');
const geminiSystemPromptInput = document.getElementById('geminiSystemPromptInput');

/* -------------------------------------------------------------------------- */
/*                               Global State                                 */
/* -------------------------------------------------------------------------- */
let historyStack = [];
let historyIndex = -1;
let configRules = [];
let recognition;
let isRecognizing = false;
let ignoreResults = false;
let currentTranscript = "";
let debounceTimer;
let mediaRecorder;
let audioChunks = [];
let groqApiKey = "";
let geminiApiKey = "";
let geminiSystemPrompt = "";
let isProcessing = false;
let shouldKeepListening = false;
let globalStream = null;
let previousTextForDiff = "";
let savedSelection = { start: 0, end: 0 };

/* -------------------------------------------------------------------------- */
/*                           Configuration Logic                              */
/* -------------------------------------------------------------------------- */
function loadConfig() {
    const savedConfig = localStorage.getItem('webspeech_config');
    const configText = savedConfig || DEFAULT_CONFIG;
    configEditor.value = configText;
    parseConfig(configText);
    
    const savedKey = localStorage.getItem('webspeech_groq_key');
    if (savedKey) {
        groqApiKeyInput.value = savedKey;
        groqApiKey = savedKey;
    }
    
    const savedGeminiKey = localStorage.getItem('webspeech_gemini_key');
    if (savedGeminiKey) {
        geminiApiKeyInput.value = savedGeminiKey;
        geminiApiKey = savedGeminiKey;
    }
    
    const savedPrompt = localStorage.getItem('webspeech_gemini_prompt');
    if (savedPrompt) {
        geminiSystemPromptInput.value = savedPrompt;
        geminiSystemPrompt = savedPrompt;
    } else {
        geminiSystemPromptInput.value = DEFAULT_SYSTEM_PROMPT;
        geminiSystemPrompt = DEFAULT_SYSTEM_PROMPT;
    }
}

function parseConfig(text) {
    configRules = [];
    let currentSection = 1; 
    
    const lines = text.split('\n');
    for (let line of lines) {
        line = line.trim();
        
        // Check for section headers: "1."
        const sectionMatch = line.match(/^(\d+)\./);
        if (sectionMatch) {
            currentSection = parseInt(sectionMatch[1], 10);
            continue;
        }
        
        if (!line || line.startsWith('#') || !line.includes('=')) continue;
        
        if (currentSection === 3) {
            // Regex Operations
            const firstEqSplit = line.split('=', 2);
            if (firstEqSplit.length < 2) {
                console.warn("Invalid Section 3 rule:", line);
                continue;
            }
            const trigger = firstEqSplit[0].trim();
            const restAfterTrigger = firstEqSplit[1];
            
            const tripleColonSplit = restAfterTrigger.split(':::', 2);
            if (tripleColonSplit.length < 2) {
                console.warn("Invalid Section 3 rule:", line);
                continue;
            }
            const matchRegexStr = tripleColonSplit[0];
            const replacement = tripleColonSplit[1];
            
            configRules.push({
                trigger: trigger,
                matchRegex: matchRegexStr,
                replacement: replacement,
                type: 3
            });
        } else {
            // Section 1 & 2
            const parts = line.split('='); 
            if (parts.length < 2) {
                console.warn(`Invalid Section ${currentSection} rule:`, line);
                continue;
            }
            const trigger = parts[0].trim();
            const replacement = parts.slice(1).join('=');
            
            configRules.push({
                trigger: trigger,
                replacement: replacement,
                type: currentSection,
                isCommand: currentSection === 1
            });
        }
    }
    console.log("Parsed config rules:", configRules);
}

/* -------------------------------------------------------------------------- */
/*                     ContentEditable Helper Functions                       */
/* -------------------------------------------------------------------------- */
function getTextContent() {
    return textBox.textContent || "";
}

function setTextContent(text) {
    textBox.textContent = text;
}

function getCursorPosition() {
    const selection = window.getSelection();
    if (!selection.rangeCount) return { start: 0, end: 0 };

    const range = selection.getRangeAt(0);
    const preRange = range.cloneRange();
    preRange.selectNodeContents(textBox);
    preRange.setEnd(range.startContainer, range.startOffset);
    const start = preRange.toString().length;

    const end = start + range.toString().length;
    return { start, end };
}

function setCursorPosition(start, end = start) {
    const selection = window.getSelection();
    const range = document.createRange();

    let currentPos = 0;
    let startNode = null, startOffset = 0;
    let endNode = null, endOffset = 0;

    function traverse(node) {
        if (startNode && endNode) return;

        if (node.nodeType === Node.TEXT_NODE) {
            const textLength = node.textContent.length;
            if (!startNode && currentPos + textLength >= start) {
                startNode = node;
                startOffset = start - currentPos;
            }
            if (!endNode && currentPos + textLength >= end) {
                endNode = node;
                endOffset = end - currentPos;
            }
            currentPos += textLength;
        } else {
            for (let child of node.childNodes) {
                traverse(child);
                if (startNode && endNode) return;
            }
        }
    }

    traverse(textBox);

    if (!startNode) {
        // If position is beyond content, place at end
        range.selectNodeContents(textBox);
        range.collapse(false);
    } else {
        range.setStart(startNode, startOffset);
        if (endNode) {
            range.setEnd(endNode, endOffset);
        } else {
            range.setEnd(startNode, startOffset);
        }
    }

    selection.removeAllRanges();
    selection.addRange(range);
}

/* -------------------------------------------------------------------------- */
/*                         Diff Highlighting Functions                        */
/* -------------------------------------------------------------------------- */
function calculateCharDiff(oldText, newText) {
    // Returns array of character indices that changed (in newText)
    // Uses a simple LCS-based diff algorithm to handle insertions/deletions
    const changed = new Set();

    // Find longest common prefix
    let prefixLen = 0;
    while (prefixLen < oldText.length && prefixLen < newText.length &&
           oldText[prefixLen] === newText[prefixLen]) {
        prefixLen++;
    }

    // Find longest common suffix (after the prefix)
    let suffixLen = 0;
    while (suffixLen < oldText.length - prefixLen &&
           suffixLen < newText.length - prefixLen &&
           oldText[oldText.length - 1 - suffixLen] === newText[newText.length - 1 - suffixLen]) {
        suffixLen++;
    }

    // Mark all characters in the middle section as changed
    const changeStart = prefixLen;
    const changeEnd = newText.length - suffixLen;

    for (let i = changeStart; i < changeEnd; i++) {
        changed.add(i);
    }

    return Array.from(changed).sort((a, b) => a - b);
}

function applyDiffHighlighting(changedIndices) {
    if (changedIndices.length === 0) return;

    const text = getTextContent();
    const { start: cursorStart, end: cursorEnd } = getCursorPosition();

    // Clear existing highlights
    textBox.querySelectorAll('.diff-highlight').forEach(span => {
        const text = span.textContent;
        span.replaceWith(document.createTextNode(text));
    });

    // Normalize text nodes
    textBox.normalize();

    // Group consecutive indices into ranges
    const ranges = [];
    let rangeStart = changedIndices[0];
    let rangeEnd = changedIndices[0];

    for (let i = 1; i < changedIndices.length; i++) {
        if (changedIndices[i] === rangeEnd + 1) {
            rangeEnd = changedIndices[i];
        } else {
            ranges.push({ start: rangeStart, end: rangeEnd + 1 });
            rangeStart = changedIndices[i];
            rangeEnd = changedIndices[i];
        }
    }
    ranges.push({ start: rangeStart, end: rangeEnd + 1 });

    // Apply highlighting from end to start to avoid offset issues
    ranges.reverse();

    for (const range of ranges) {
        let currentPos = 0;
        let applied = false;

        function highlightInNode(node) {
            if (applied) return;

            if (node.nodeType === Node.TEXT_NODE) {
                const textLength = node.textContent.length;
                const nodeStart = currentPos;
                const nodeEnd = currentPos + textLength;

                // Check if this node contains part of our range
                if (range.start < nodeEnd && range.end > nodeStart) {
                    const highlightStart = Math.max(0, range.start - nodeStart);
                    const highlightEnd = Math.min(textLength, range.end - nodeStart);

                    const before = node.textContent.substring(0, highlightStart);
                    const highlighted = node.textContent.substring(highlightStart, highlightEnd);
                    const after = node.textContent.substring(highlightEnd);

                    const fragment = document.createDocumentFragment();
                    if (before) fragment.appendChild(document.createTextNode(before));

                    const span = document.createElement('span');
                    span.className = 'diff-highlight';
                    span.textContent = highlighted;
                    fragment.appendChild(span);

                    if (after) fragment.appendChild(document.createTextNode(after));

                    node.replaceWith(fragment);
                    applied = true;
                    return;
                }

                currentPos += textLength;
            } else {
                const children = Array.from(node.childNodes);
                for (let child of children) {
                    highlightInNode(child);
                    if (applied) return;
                }
            }
        }

        highlightInNode(textBox);
        currentPos = 0;
        applied = false;
    }

    // Restore cursor position
    setCursorPosition(cursorStart, cursorEnd);
}

function clearDiffHighlighting() {
    textBox.querySelectorAll('.diff-highlight').forEach(span => {
        const text = span.textContent;
        span.replaceWith(document.createTextNode(text));
    });
    textBox.normalize();
}

function trackAndHighlightChanges() {
    const currentText = getTextContent();
    if (previousTextForDiff !== currentText) {
        const changedIndices = calculateCharDiff(previousTextForDiff, currentText);
        clearDiffHighlighting();
        applyDiffHighlighting(changedIndices);
        previousTextForDiff = currentText;
    }
}

/* -------------------------------------------------------------------------- */
/*                             network Execution                              */
/* -------------------------------------------------------------------------- */
async function executeWithGemini(instruction) {
    statusDiv.textContent = `Status: Waiting for Gemini to process: "${instruction}"`;
    if (!geminiApiKey) {
        alert("Please set your Gemini API Key in configuration.");
        statusDiv.textContent = "Status: Missing Gemini Key";
        return;
    }

    // Clear previous highlighting and track text before change
    clearDiffHighlighting();

    // Remove all cursor symbols for clean processing
    let currentText = getTextContent().replace(new RegExp(CURSOR_SYMBOL, 'g'), '');
    previousTextForDiff = currentText;

    // Prepare context with markers
    let { start: selStart, end: selEnd } = savedSelection;

    // Check if cursor symbol was in the original text
    const originalText = getTextContent();
    if (originalText.includes(CURSOR_SYMBOL)) {
        const idx = originalText.indexOf(CURSOR_SYMBOL);
        selStart = idx;
        selEnd = idx;
    }

    // Calculate cursor line index for fallback
    const textBeforeCursor = currentText.substring(0, selStart);
    const cursorLineIndex = (textBeforeCursor.match(/\n/g) || []).length;

    const markedText = currentText.slice(0, selStart) + MARKER_A + currentText.slice(selStart, selEnd) + MARKER_B + currentText.slice(selEnd);

    const payload = {
        "contents": [{
            "parts": [{"text": markedText}]
        }],
        "system_instruction": {
            "parts": [{ "text": geminiSystemPrompt + instruction}]
        }
    };

    try {
        const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${geminiApiKey}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });

        if (!response.ok) {
            const errText = await response.text();
            throw new Error(`Gemini API Error: ${response.status} - ${errText}`);
        }

        const data = await response.json();
        if (data.candidates && data.candidates.length > 0 && data.candidates[0].content && data.candidates[0].content.parts) {
            let newText = data.candidates[0].content.parts[0].text;
            console.log("Gemini Response:", newText);

            saveState();

            const aIdx = newText.indexOf(MARKER_A);
            const bIdx = newText.indexOf(MARKER_B);

            if (aIdx !== -1 || bIdx !== -1) {
                const cleanText = newText.replace(new RegExp(MARKER_A, 'g'), '').replace(new RegExp(MARKER_B, 'g'), '');
                setTextContent(cleanText);

                let start = aIdx;
                let end = bIdx;

                if (start !== -1 && end !== -1) {
                    if (start < end) {
                        end -= MARKER_A.length;
                    } else {
                        start -= MARKER_B.length;
                    }
                    setCursorPosition(start, end);
                } else if (start !== -1) {
                    setCursorPosition(start, start);
                } else if (end !== -1) {
                    setCursorPosition(end, end);
                }
            } else {
                // Fallback: No markers found. Restore cursor to original line index.
                setTextContent(newText);

                const lines = newText.split('\n');
                let targetLine = cursorLineIndex;

                if (targetLine >= lines.length) {
                    targetLine = lines.length - 1;
                }
                if (targetLine < 0) targetLine = 0;

                let charIndex = 0;
                for (let i = 0; i < targetLine; i++) {
                    charIndex += lines[i].length + 1; // +1 for newline char
                }

                setCursorPosition(charIndex, charIndex);
            }

            saveState();
            trackAndHighlightChanges();
            scrollToCursor();
            statusDiv.textContent = "Status: Execution Complete";
        }
    } catch (e) {
        console.error("Gemini Execute Failed:", e);
        statusDiv.textContent = "Status: Gemini Error";
        alert("Gemini Execution Failed: " + e.message);
    }
}


async function processWithGroq(mode = 'process') {
    if (isProcessing) return;
    isProcessing = true;
    statusDiv.textContent = "Status: Processing...";
    
    const wasListening = shouldKeepListening;
    shouldKeepListening = false;
    
    try {
        // 1. Stop Recording to finalize chunks
        if (isRecognizing || (mediaRecorder && mediaRecorder.state === "recording")) {
            await new Promise(resolve => {
                if (!mediaRecorder || mediaRecorder.state === "inactive") {
                    resolve();
                    return;
                }
                mediaRecorder.onstop = () => resolve();
                mediaRecorder.stop();
            });
            
            if (isRecognizing) recognition.stop();
        }
        
        // 2. Prepare Audio Blob
        if (groqApiKey && audioChunks.length > 0) {
            statusDiv.textContent = "Status: Sending to Groq...";
            const blob = new Blob(audioChunks, { type: 'audio/webm' });
            
            // 3. Send to Groq
            try {
                const formData = new FormData();
                formData.append('file', blob, 'recording.webm');
                formData.append('model', 'whisper-large-v3-turbo');
                
                const response = await fetch('https://api.groq.com/openai/v1/audio/transcriptions', {
                    method: 'POST',
                    headers: {
                        'Authorization': `Bearer ${groqApiKey}`
                    },
                    body: formData
                });
                
                if (!response.ok) {
                    const errText = await response.text();
                    throw new Error(`Groq API Error: ${response.status} - ${errText}`);
                }
                
                const data = await response.json();
                let groqText = data.text;
                console.log("Groq Transcription Raw:", groqText);
                
                let isExecuteCommand = mode === 'execute';
                
                // Clean command words
                const processRules = configRules.filter(r => r.type === 1 && ['#process', '#execute'].includes(r.replacement));
                
                for (const rule of processRules) {
                    try {
                        const regex = new RegExp(`(?:^|\\s)(${rule.trigger})[\\s.,!?]*$`, 'i');
                        if (regex.test(groqText)) {
                            if (rule.replacement === '#execute') isExecuteCommand = true;
                            groqText = groqText.replace(regex, "").trim();
                        }
                    } catch (e) { } // Ignore regex errors
                }
                
                console.log("Groq Transcription Cleaned:", groqText, "Mode:", isExecuteCommand ? "Execute" : "Process");
                
                if (isExecuteCommand) {
                    await executeWithGemini(groqText);
                } else {
                    runTextProcessing(groqText);
                }
                
                if (!statusDiv.textContent.includes("Error")) {
                    statusDiv.textContent = "Status: Ready";
                }
                
            } catch (e) {
                console.error("Groq Processing Failed:", e);
                statusDiv.textContent = "Status: Groq Error (Falling back to WebSpeech)";
                // Fallback
                runTextProcessing(currentTranscript);
            }
        } else {
            if (!groqApiKey) console.log("No Groq API Key provided, using WebSpeech.");
            else console.log("No audio chunks captured.");
            runTextProcessing(currentTranscript);
            statusDiv.textContent = "Status: Ready";
        }
    } finally {
        // Cleanup
        audioChunks = [];
        currentTranscript = "";
        pendingTextSpan.textContent = "";
        ignoreResults = true; 
        isProcessing = false;
        
        if (wasListening) {
            startDictation();
        }
    }
}

/* -------------------------------------------------------------------------- */
/*                           Audio Processing Logic                           */
/* -------------------------------------------------------------------------- */

function restartForNewSegment() {
    // Flags to stop processing current events but keep listening intent
    ignoreResults = true; 
    
    // Stop current capture - onend will handle restart because shouldKeepListening is true
    if (mediaRecorder && mediaRecorder.state !== 'inactive') {
        mediaRecorder.stop();
    }
    if (isRecognizing) {
        recognition.stop(); 
    }
}

function discardText() {
    console.log("Discarding text");
    currentTranscript = "";
    pendingTextSpan.textContent = "";
    restartForNewSegment();
}

function stopDictation() {
    shouldKeepListening = false;
    if (mediaRecorder && mediaRecorder.state !== "inactive") {
        mediaRecorder.stop();
    }
    recognition.stop();

    // Stop all microphone tracks to fully release the microphone
    if (globalStream) {
        globalStream.getTracks().forEach(track => track.stop());
        globalStream = null;
    }
}

async function startDictation() {
    shouldKeepListening = true;
    audioChunks = [];
    
    // Init MediaRecorder for Groq
    try {
        if (!globalStream) {
            globalStream = await navigator.mediaDevices.getUserMedia({ audio: true });
        }
        
        mediaRecorder = new MediaRecorder(globalStream);
        mediaRecorder.ondataavailable = (event) => {
            if (event.data.size > 0) {
                audioChunks.push(event.data);
            }
        };
        mediaRecorder.start();
    } catch (err) {
        console.error("Error accessing microphone:", err);
        statusDiv.textContent = "Status: Mic Error (Check permissions)";
    }
    
    try {
        recognition.start();
    } catch(e) {
        // Already started?
    }
}

/* -------------------------------------------------------------------------- */
/*                         Speech Recognition Setup                           */
/* -------------------------------------------------------------------------- */
if (window.SpeechRecognition || window.webkitSpeechRecognition) {
    const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
    recognition = new SpeechRecognition();
    recognition.continuous = true;
    recognition.interimResults = true;
    recognition.lang = 'en-US';
    
    recognition.onstart = () => {
        isRecognizing = true;
        ignoreResults = false;
        currentTranscript = "";
        pendingTextSpan.textContent = "";
        toggleButton.textContent = 'Stop Dictation';
        toggleButton.classList.add('recording');
        statusDiv.textContent = "Status: Listening...";
    };
    
    recognition.onerror = (event) => {
        if (event.error === 'no-speech') {
            console.log("No speech detected, recognition will continue.");
            return;
        }
        console.error("Speech recognition error", event.error);
        statusDiv.textContent = "Status: Error - " + event.error;
        isRecognizing = false;
        toggleButton.textContent = 'Start Dictation';
        toggleButton.classList.remove('recording');
    };
    
    recognition.onend = () => {
        // If we are not purposefully processing or stopping, restart
        if (shouldKeepListening) { 
            startDictation();
        } else {
            // Only update UI if we are NOT processing (to prevent flickering)
            if (!isProcessing) {
                isRecognizing = false;
                toggleButton.textContent = 'Start Dictation';
                toggleButton.classList.remove('recording');
                if (!statusDiv.textContent.includes("Processing")) {
                    statusDiv.textContent = "Status: Stopped";
                }
            }
        }
    };
    
    recognition.onresult = (event) => {
        if (ignoreResults) return;
        
        let final = "";
        for (let i = event.resultIndex; i < event.results.length; ++i) {
            if (event.results[i].isFinal) {
                final += event.results[i][0].transcript;
            }
        }
        
        const fullTranscript = Array.from(event.results)
        .map(res => res[0].transcript)
        .join('');
        
        currentTranscript = fullTranscript;
        const trimmedTranscript = currentTranscript.trim();
        
        // 1. Find triggers for #process and #discard from config
        let processTriggerRegex = null;
        let executeTriggerRegex = null;
        let discardTriggerRegex = null;
        
        for (const rule of configRules) {
            if (rule.type === 1) { // Command
                try {
                    const regex = new RegExp(`(?:^|\\s)(${rule.trigger})\\s*$`, 'i');
                    if (regex.test(currentTranscript)) {
                        if (rule.replacement === '#process') processTriggerRegex = regex;
                        if (rule.replacement === '#execute') executeTriggerRegex = regex;
                        if (rule.replacement === '#discard') discardTriggerRegex = regex;
                    }
                } catch (e) {}
            }
        }
        
        // 2. Standard Auto-match (only if not process/execute/discard)
        let autoMatchedRule = null;
        if (!processTriggerRegex && !executeTriggerRegex && !discardTriggerRegex) {
            for (const rule of configRules) {
                if (rule.type === 1 || rule.type === 2 || rule.type === 3) {
                    try {
                        const regex = new RegExp(`^(${rule.trigger})$`, 'i');
                        if (regex.test(trimmedTranscript)) {
                            autoMatchedRule = rule;
                            break;
                        }
                    } catch (e) {}
                }
            }
        }
        
        if (processTriggerRegex) {
            const match = currentTranscript.match(processTriggerRegex);
            if (match) {
                currentTranscript = currentTranscript.substring(0, match.index);
                processWithGroq('process');
            }
        } else if (executeTriggerRegex) {
            const match = currentTranscript.match(executeTriggerRegex);
            if (match) {
                currentTranscript = currentTranscript.substring(0, match.index);
                processWithGroq('execute');
            }
        } else if (discardTriggerRegex) {
            discardText();
        } else if (autoMatchedRule) {
            // Standard commands execute immediately on local transcript
            currentTranscript = "";
            pendingTextSpan.textContent = "";
            runTextProcessing(trimmedTranscript);
            restartForNewSegment();
        } else {
            pendingTextSpan.textContent = currentTranscript;
        }
    };
    
} else {
    toggleButton.disabled = true;
    processButton.disabled = true;
    statusDiv.textContent = 'Speech recognition is not supported in this browser.';
}



/* -------------------------------------------------------------------------- */
/*                             History Management                             */
/* -------------------------------------------------------------------------- */
function saveState() {
    // Clean symbol before saving
    const currentVal = getTextContent().replace(new RegExp(CURSOR_SYMBOL, 'g'), '');
    localStorage.setItem('webspeech_content', currentVal);

    if (historyIndex >= 0 && historyStack[historyIndex] === currentVal) {
        return;
    }

    if (historyIndex < historyStack.length - 1) {
        historyStack = historyStack.slice(0, historyIndex + 1);
    }
    historyStack.push(currentVal);
    if (historyStack.length > MAX_HISTORY) {
        historyStack.shift();
    } else {
        historyIndex++;
    }
}

function restoreState() {
    const savedContent = localStorage.getItem('webspeech_content');
    if (savedContent) {
        setTextContent(savedContent.replace(new RegExp(CURSOR_SYMBOL, 'g'), ''));
        previousTextForDiff = getTextContent();
    }
}

function undo() {
    if (historyIndex > 0) {
        clearDiffHighlighting();
        // Clean text of cursor symbols for diff tracking
        previousTextForDiff = getTextContent().replace(new RegExp(CURSOR_SYMBOL, 'g'), '');
        historyIndex--;
        setTextContent(historyStack[historyIndex]);
        localStorage.setItem('webspeech_content', getTextContent());
        trackAndHighlightChanges();
        statusDiv.textContent = "Undo";
    } else {
        statusDiv.textContent = "Nothing to undo";
    }
}

function redo() {
    if (historyIndex < historyStack.length - 1) {
        clearDiffHighlighting();
        // Clean text of cursor symbols for diff tracking
        previousTextForDiff = getTextContent().replace(new RegExp(CURSOR_SYMBOL, 'g'), '');
        historyIndex++;
        setTextContent(historyStack[historyIndex]);
        localStorage.setItem('webspeech_content', getTextContent());
        trackAndHighlightChanges();
        statusDiv.textContent = "Redo";
    } else {
        statusDiv.textContent = "Nothing to redo";
    }
}

function scrollToCursor() {
    // Wait for value update
    setTimeout(() => {
        const selection = window.getSelection();
        if (!selection.rangeCount) return;

        const range = selection.getRangeAt(0);
        const rect = range.getBoundingClientRect();
        const containerRect = textBox.getBoundingClientRect();

        // Calculate scroll position relative to container
        const relativeTop = rect.top - containerRect.top + textBox.scrollTop;

        // Scroll logic
        const currentScrollTop = textBox.scrollTop;
        const clientHeight = textBox.clientHeight;

        // If cursor is below view
        if (relativeTop + rect.height > currentScrollTop + clientHeight) {
            textBox.scrollTop = relativeTop + rect.height - clientHeight + 30;
        }
        // If cursor is above view
        else if (relativeTop < currentScrollTop) {
            textBox.scrollTop = relativeTop - 30;
        }

    }, 0);
}

/* -------------------------------------------------------------------------- */
/*                            Cursor & Text Logic                             */
/* -------------------------------------------------------------------------- */
function insertTextAtCursor(text) {
    saveState();

    // Clear previous highlighting and track text before change
    clearDiffHighlighting();

    // Remove all cursor symbols for clean processing
    let currentText = getTextContent().replace(new RegExp(CURSOR_SYMBOL, 'g'), '');
    previousTextForDiff = currentText;

    let { start: startPos, end: endPos } = savedSelection;

    // Check if cursor symbol was in the original text
    const originalText = getTextContent();
    const cursorSymbolIndex = originalText.indexOf(CURSOR_SYMBOL);
    if (cursorSymbolIndex !== -1) {
        startPos = cursorSymbolIndex;
        endPos = cursorSymbolIndex;
    }

    const textBefore = currentText.slice(0, startPos);
    const textAfter = currentText.slice(endPos);
    const processedText = text.replace(/\\n/g, '\n');

    setTextContent(textBefore + processedText + textAfter);

    const newCursorPos = startPos + processedText.length;
    setCursorPosition(newCursorPos, newCursorPos);

    saveState();
    trackAndHighlightChanges();

    // Re-add cursor symbol if textbox is not focused
    if (document.activeElement !== textBox) {
        const val = getTextContent();
        setTextContent(val.slice(0, newCursorPos) + CURSOR_SYMBOL + val.slice(newCursorPos));
    }

    scrollToCursor();
}

/* -------------------------------------------------------------------------- */
/*                             Command Registry                               */
/* -------------------------------------------------------------------------- */
const commandRegistry = {
    '#undo': undo,
    '#redo': redo,
    '#stop': stopDictation,
    '#discard': discardText,
};

function runTextProcessing(rawTextInput) {
    const rawText = rawTextInput ? rawTextInput.trim() : "";
    if (!rawText) return;
    
    console.log("Processing:", rawText);
    let actionTaken = false;
    let matchedRule = null;
    
    // Find matching rule
    for (const rule of configRules) {
        try {
            const regex = new RegExp(`^(${rule.trigger})$`, 'i');
            if (regex.test(rawText)) {
                matchedRule = rule;
                break;
            }
        } catch (e) {
            console.warn("Bad regex in rule:", rule.trigger);
        }
    }
    
    if (matchedRule) {
        if (matchedRule.type === 1) { // Command
            const cmdFunc = commandRegistry[matchedRule.replacement];
            if (cmdFunc) {
                if (!['#undo', '#redo', '#stop', '#discard'].includes(matchedRule.replacement)) {
                    saveState(); 
                }
                cmdFunc();
                if (!['#undo', '#redo', '#stop', '#discard'].includes(matchedRule.replacement)) {
                    saveState(); 
                }
                actionTaken = true;
            } else {
                console.warn("Unknown command:", matchedRule.replacement);
            }
        } else if (matchedRule.type === 3) { // Regex Operation
            saveState();

            // Clear previous highlighting and track text before change
            clearDiffHighlighting();

            // Remove all cursor symbols for clean processing
            let currentText = getTextContent().replace(new RegExp(CURSOR_SYMBOL, 'g'), '');
            previousTextForDiff = currentText;

            let { start: selStart, end: selEnd } = savedSelection;

            // Check if cursor symbol was in the original text
            const originalText = getTextContent();
            if (originalText.includes(CURSOR_SYMBOL)) {
                const idx = originalText.indexOf(CURSOR_SYMBOL);
                selStart = idx;
                selEnd = idx;
            }

            // Construct marked text
            const markedText = currentText.slice(0, selStart) + MARKER_A + currentText.slice(selStart, selEnd) + MARKER_B + currentText.slice(selEnd);

            try {
                const opRegex = new RegExp(matchedRule.matchRegex, 'gm');
                let replacementStr = matchedRule.replacement.replace(/\\n/g, '\n');
                const newMarkedText = markedText.replace(opRegex, replacementStr);

                const newAIndex = newMarkedText.indexOf(MARKER_A);
                const newBIndex = newMarkedText.indexOf(MARKER_B);

                let finalCleanText = newMarkedText.replace(new RegExp(MARKER_A, 'g'), '').replace(new RegExp(MARKER_B, 'g'), '');
                setTextContent(finalCleanText);

                if (newAIndex !== -1 && newBIndex !== -1) {
                    let newStart = newAIndex;
                    let newEnd = newBIndex;
                    if (newStart < newEnd) {
                        newEnd -= MARKER_A.length;
                    } else {
                        newStart -= MARKER_B.length;
                    }
                    setCursorPosition(newStart, newEnd);
                } else if (newAIndex !== -1) {
                    setCursorPosition(newAIndex, newAIndex);
                }

                actionTaken = true;
                saveState();
                trackAndHighlightChanges();
                scrollToCursor();
            } catch (e) {
                console.error("Regex Op Failed", e);
            }
        } else { // Substitution
            insertTextAtCursor(matchedRule.replacement);
            actionTaken = true;
        }
    } else {
        // Append Logic
        const currentVal = getTextContent();
        let insertionPos = savedSelection.start;
        if (currentVal.includes(CURSOR_SYMBOL)) {
            insertionPos = currentVal.indexOf(CURSOR_SYMBOL);
        }

        // Don't trim newlines, only spaces and tabs for proper sentence/line detection
        const textBefore = currentVal.substring(0, insertionPos).replace(/[ \t]+$/, '');
        const lastChar = textBefore.length > 0 ? textBefore.charAt(textBefore.length - 1) : "";
        const isSentenceStart = textBefore.length === 0 || [".", "!", "?", "\n"].includes(lastChar);

        // Check if there's any alphabetic character on the current line
        const lastNewlineIndex = textBefore.lastIndexOf('\n');
        const currentLineText = lastNewlineIndex >= 0
            ? textBefore.substring(lastNewlineIndex + 1)
            : textBefore;
        const hasAlphaOnLine = /[a-zA-Z]/.test(currentLineText);

        let textToAppend = rawText;
        if (textToAppend.length > 0) {
            // Capitalize if at sentence start OR if current line has no alphabetic characters yet
            const shouldCapitalize = isSentenceStart || !hasAlphaOnLine;

            // Find the first alphabetic character
            const alphaMatch = textToAppend.match(/[a-zA-Z]/);
            if (alphaMatch) {
                const index = alphaMatch.index;
                if (shouldCapitalize) {
                    textToAppend = textToAppend.substring(0, index) +
                                  textToAppend.charAt(index).toUpperCase() +
                                  textToAppend.substring(index + 1);
                } else {
                    textToAppend = textToAppend.substring(0, index) +
                                  textToAppend.charAt(index).toLowerCase() +
                                  textToAppend.substring(index + 1);
                }
            }
        }

        const charBeforeCursor = currentVal.length > 0 && insertionPos > 0 ? currentVal.charAt(insertionPos - 1) : "";
        const needsSpace = charBeforeCursor && !["\n", " "].includes(charBeforeCursor);

        insertTextAtCursor((needsSpace ? " " : "") + textToAppend);
        actionTaken = true;
    }
}

/* -------------------------------------------------------------------------- */
/*                             UI Event Listeners                             */
/* -------------------------------------------------------------------------- */
// Sidebar Logic
sidebarToggle.addEventListener('click', () => {
    sidebar.classList.toggle('collapsed');
    localStorage.setItem('webspeech_sidebar_collapsed', sidebar.classList.contains('collapsed'));
});

// Restore Sidebar State
if (localStorage.getItem('webspeech_sidebar_collapsed') === 'true') {
    sidebar.classList.add('collapsed');
} else {
    sidebar.classList.remove('collapsed');
}

saveConfigButton.addEventListener('click', () => {
    const text = configEditor.value;
    localStorage.setItem('webspeech_config', text);
    parseConfig(text);
    
    const key = groqApiKeyInput.value.trim();
    groqApiKey = key;
    localStorage.setItem('webspeech_groq_key', key);
    
    const geminiKey = geminiApiKeyInput.value.trim();
    geminiApiKey = geminiKey;
    localStorage.setItem('webspeech_gemini_key', geminiKey);
    
    const prompt = geminiSystemPromptInput.value.trim();
    geminiSystemPrompt = prompt;
    localStorage.setItem('webspeech_gemini_prompt', prompt);
    
    alert("Configuration & Keys saved!");
});

resetConfigButton.addEventListener('click', () => {
    if (confirm("Are you sure you want to reset configuration to defaults?")) {
        configEditor.value = DEFAULT_CONFIG;
        geminiSystemPromptInput.value = DEFAULT_SYSTEM_PROMPT;
        localStorage.removeItem('webspeech_config'); 
        localStorage.removeItem('webspeech_gemini_prompt');
        parseConfig(DEFAULT_CONFIG);
        geminiSystemPrompt = DEFAULT_SYSTEM_PROMPT;
        alert("Configuration reset to defaults. (Not saved to storage until you click Save)");
    }
});

toggleButton.addEventListener('click', () => {
    if (isRecognizing) {
        stopDictation();
    } else {
        startDictation();
    }
});

processButton.addEventListener('click', () => processWithGroq('process'));
executeButton.addEventListener('click', () => processWithGroq('execute'));
undoButton.addEventListener('click', undo);
redoButton.addEventListener('click', redo);
discardButton.addEventListener('click', discardText);

// Keyboard Control
document.addEventListener('keydown', (event) => {
    if (event.repeat) return;
    const activeElement = document.activeElement;
    const isInputFieldActive = (
        activeElement === textBox ||
        activeElement === groqApiKeyInput ||
        activeElement === geminiApiKeyInput ||
        activeElement === geminiSystemPromptInput ||
        activeElement === configEditor
    );

    // Escape key: Stop dictation
    if (event.key === 'Escape') {
        if (isRecognizing) {
            event.preventDefault();
            stopDictation();
        }
        return;
    }

    if (event.key === 'Control' || (event.key === ' ' && !isInputFieldActive)) {
        if (event.key === ' ') {
            event.preventDefault();
        }
        if (!isRecognizing) {
            startDictation();
        } else {
            processWithGroq('process');
        }
    }
});

document.addEventListener('selectionchange', () => {
    const selection = window.getSelection();
    if (selection.rangeCount > 0 && textBox.contains(selection.getRangeAt(0).commonAncestorContainer)) {
        savedSelection = getCursorPosition();
    }
});

// Textbox Events (Symbol handling & State)
textBox.addEventListener('blur', () => {
    const val = getTextContent();
    // Only add cursor symbol if it doesn't already exist and selection is collapsed
    if (!val.includes(CURSOR_SYMBOL)) {
        if (savedSelection.start === savedSelection.end) {
            setTextContent(val.slice(0, savedSelection.start) + CURSOR_SYMBOL + val.slice(savedSelection.start));
        }
    }
});

textBox.addEventListener('focus', () => {
    const val = getTextContent();
    const pos = val.indexOf(CURSOR_SYMBOL);
    if (pos !== -1) {
        setTextContent(val.replace(new RegExp(CURSOR_SYMBOL, 'g'), ''));
        setCursorPosition(pos, pos);
    }
});

textBox.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
        e.preventDefault();
        const selection = window.getSelection();
        if (selection.rangeCount) {
            const range = selection.getRangeAt(0);
            range.deleteContents();
            const textNode = document.createTextNode('\n');
            range.insertNode(textNode);
            range.setStartAfter(textNode);
            range.setEndAfter(textNode);
            selection.removeAllRanges();
            selection.addRange(range);
            textBox.dispatchEvent(new Event('input'));
        }
    }
});

textBox.addEventListener('paste', (e) => {
    e.preventDefault();
    const text = (e.clipboardData || window.clipboardData).getData('text');
    document.execCommand('insertText', false, text);
});

textBox.addEventListener('input', () => {
    clearTimeout(debounceTimer);
    debounceTimer = setTimeout(() => {
        saveState();
    }, 1000);
});

// Initialization
loadConfig();
restoreState();
saveState(); // Initial stack push
