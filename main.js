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
/*                             network Execution                              */
/* -------------------------------------------------------------------------- */
async function executeWithGemini(instruction) {
    statusDiv.textContent = `Status: Waiting for Gemini to process: "${instruction}"`;
    if (!geminiApiKey) {
        alert("Please set your Gemini API Key in configuration.");
        statusDiv.textContent = "Status: Missing Gemini Key";
        return;
    }
    
    // Prepare context with markers
    let currentText = textBox.value;
    let selStart = textBox.selectionStart;
    let selEnd = textBox.selectionEnd;
    
    if (currentText.includes(CURSOR_SYMBOL)) {
        const idx = currentText.indexOf(CURSOR_SYMBOL);
        selStart = idx;
        selEnd = idx;
        currentText = currentText.replace(new RegExp(CURSOR_SYMBOL, 'g'), '');
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
            
            // Use trimmed text for consistency
            newText = newText.trim();
            
            const aIdx = newText.indexOf(MARKER_A);
            const bIdx = newText.indexOf(MARKER_B);
            
            if (aIdx !== -1 || bIdx !== -1) {
                const cleanText = newText.replace(new RegExp(MARKER_A, 'g'), '').replace(new RegExp(MARKER_B, 'g'), '');
                textBox.value = cleanText;
                
                let start = aIdx;
                let end = bIdx;
                
                if (start !== -1 && end !== -1) {
                    if (start < end) {
                        end -= MARKER_A.length; 
                    } else {
                        start -= MARKER_B.length;
                    }
                    textBox.setSelectionRange(start, end);
                } else if (start !== -1) {
                    textBox.setSelectionRange(start, start);
                } else if (end !== -1) {
                    // If only B is present (unlikely), treat as cursor pos
                    // We need to adjust for A removal if A was somehow before it? No, A is not present.
                    textBox.setSelectionRange(end, end);
                }
            } else {
                // Fallback: No markers found. Restore cursor to original line index.
                textBox.value = newText;
                
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
                
                textBox.setSelectionRange(charIndex, charIndex);
            }
            
            saveState();
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
    const currentVal = textBox.value.replace(new RegExp(CURSOR_SYMBOL, 'g'), '');
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
        textBox.value = savedContent.replace(new RegExp(CURSOR_SYMBOL, 'g'), '');
    }
}

function undo() {
    if (historyIndex > 0) {
        historyIndex--;
        textBox.value = historyStack[historyIndex];
        localStorage.setItem('webspeech_content', textBox.value);
        statusDiv.textContent = "Undo";
    } else {
        statusDiv.textContent = "Nothing to undo";
    }
}

function redo() {
    if (historyIndex < historyStack.length - 1) {
        historyIndex++;
        textBox.value = historyStack[historyIndex];
        localStorage.setItem('webspeech_content', textBox.value);
        statusDiv.textContent = "Redo";
    } else {
        statusDiv.textContent = "Nothing to redo";
    }
}

function scrollToCursor() {
    // Wait for value update
    setTimeout(() => {
        const val = textBox.value;
        const sel = textBox.selectionStart;
        
        // Create a dummy div to mirror the textarea's properties
        const div = document.createElement('div');
        const style = window.getComputedStyle(textBox);
        
        // Copy all relevant styles
        Array.from(style).forEach(prop => {
            div.style[prop] = style.getPropertyValue(prop);
        });
        
        div.style.position = 'absolute';
        div.style.top = '-9999px';
        div.style.left = '-9999px';
        div.style.height = 'auto';
        
        // Important: Use clientWidth to match the visible area (excluding scrollbar)
        // and force border-box to ensure padding is handled consistently.
        div.style.boxSizing = 'border-box';
        div.style.width = textBox.clientWidth + 'px';
        
        div.style.overflow = 'hidden';
        div.style.whiteSpace = 'pre-wrap'; 
        div.style.wordWrap = 'break-word';
        
        // The text before the cursor
        const content = val.substring(0, sel);
        // Use a span to mark the cursor position
        const span = document.createElement('span');
        span.textContent = '|';
        
        div.textContent = content;
        div.appendChild(span);
        
        document.body.appendChild(div);
        
        const spanTop = span.offsetTop;
        const spanHeight = span.offsetHeight;
        
        // Cleanup
        document.body.removeChild(div);
        
        // Scroll logic
        const currentScrollTop = textBox.scrollTop;
        const clientHeight = textBox.clientHeight;
        
        // If cursor is below view
        if (spanTop + spanHeight > currentScrollTop + clientHeight) {
            // Scroll so the bottom of cursor is visible, plus some buffer
            textBox.scrollTop = spanTop + spanHeight - clientHeight + 30; 
        } 
        // If cursor is above view
        else if (spanTop < currentScrollTop) {
            textBox.scrollTop = spanTop - 30;
        }
        
    }, 0);
}

/* -------------------------------------------------------------------------- */
/*                            Cursor & Text Logic                             */
/* -------------------------------------------------------------------------- */
function insertTextAtCursor(text) {
    saveState();
    
    let currentText = textBox.value;
    let startPos = textBox.selectionStart;
    let endPos = textBox.selectionEnd;
    
    const cursorSymbolIndex = currentText.indexOf(CURSOR_SYMBOL);
    if (cursorSymbolIndex !== -1) {
        startPos = cursorSymbolIndex;
        endPos = cursorSymbolIndex;
        currentText = currentText.replace(CURSOR_SYMBOL, '');
    }
    
    const textBefore = currentText.slice(0, startPos);
    const textAfter = currentText.slice(endPos);
    const processedText = text.replace(/\\n/g, '\n');
    
    textBox.value = textBefore + processedText + textAfter;
    
    const newCursorPos = startPos + processedText.length;
    textBox.setSelectionRange(newCursorPos, newCursorPos); 
    
    if (cursorSymbolIndex !== -1) {
        const val = textBox.value;
        textBox.value = val.slice(0, newCursorPos) + CURSOR_SYMBOL + val.slice(newCursorPos);
    }
    
    saveState(); 
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
            let currentText = textBox.value;
            let selStart = textBox.selectionStart;
            let selEnd = textBox.selectionEnd;
            
            if (currentText.includes(CURSOR_SYMBOL)) {
                const idx = currentText.indexOf(CURSOR_SYMBOL);
                selStart = idx;
                selEnd = idx;
                currentText = currentText.replace(CURSOR_SYMBOL, '');
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
                textBox.value = finalCleanText;
                
                if (newAIndex !== -1 && newBIndex !== -1) {
                    let newStart = newAIndex;
                    let newEnd = newBIndex;
                    if (newStart < newEnd) {
                        newEnd -= MARKER_A.length;
                    } else {
                        newStart -= MARKER_B.length;
                    }
                    textBox.setSelectionRange(newStart, newEnd);
                } else if (newAIndex !== -1) {
                    textBox.setSelectionRange(newAIndex, newAIndex);
                }
                
                actionTaken = true;
                saveState();
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
        const currentVal = textBox.value;
        let insertionPos = textBox.selectionStart;
        if (currentVal.includes(CURSOR_SYMBOL)) {
            insertionPos = currentVal.indexOf(CURSOR_SYMBOL);
        }
        
        const textBefore = currentVal.substring(0, insertionPos).trimEnd();
        const lastChar = textBefore.length > 0 ? textBefore.charAt(textBefore.length - 1) : "";
        const isSentenceStart = textBefore.length === 0 || [".", "!", "?", "\n"].includes(lastChar);
        
        let textToAppend = rawText;
        if (textToAppend.length > 0) {
            if (isSentenceStart) {
                textToAppend = textToAppend.charAt(0).toUpperCase() + textToAppend.slice(1);
            } else {
                textToAppend = textToAppend.charAt(0).toLowerCase() + textToAppend.slice(1);
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

// Textbox Events (Symbol handling & State)
textBox.addEventListener('blur', () => {
    const val = textBox.value;
    const sel = textBox.selectionStart;
    textBox.value = val.slice(0, sel) + CURSOR_SYMBOL + val.slice(sel);
});

textBox.addEventListener('focus', () => {
    const val = textBox.value;
    const pos = val.indexOf(CURSOR_SYMBOL);
    if (pos !== -1) {
        textBox.value = val.replace(new RegExp(CURSOR_SYMBOL, 'g'), '');
        textBox.setSelectionRange(pos, pos);
    }
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
