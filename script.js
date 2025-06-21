(function() {
    'use strict';

    class PoemCompiler {
        constructor() {
            this.poems = [];
            this.selectedFiles = [];
            this.draggedIndex = null;
            this.isProcessing = false;
            this.notificationTimeout = null;
            this.initializeEventListeners();
            this.updateDisplay();
        }

        /**
         * Initializes all event listeners for the UI elements.
         */
        initializeEventListeners() {
            const wordFiles = document.getElementById('wordFiles');
            const processBtn = document.getElementById('processBtn');
            const downloadBtn = document.getElementById('downloadBtn');
            const clearBtn = document.getElementById('clearBtn');
            const fileLabel = document.getElementById('fileLabel');

            if (!wordFiles || !processBtn || !downloadBtn || !clearBtn || !fileLabel) {
                console.error('Required DOM elements not found. Please ensure all IDs are correct in the HTML.');
                return;
            }

            // File input change event
            wordFiles.addEventListener('change', (e) => {
                this.handleFileSelect(e);
            });

            // Process button click event
            processBtn.addEventListener('click', () => {
                if (!this.isProcessing) {
                    this.processDocuments();
                }
            });

            // Download button click event
            downloadBtn.addEventListener('click', () => {
                this.downloadCombinedDocument();
            });

            // Clear button click event
            clearBtn.addEventListener('click', () => {
                this.clearAllPoems();
            });

            // --- Drag and drop functionality for the file label ---
            // Prevent default drag behaviors
            ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
                fileLabel.addEventListener(eventName, (e) => this.preventDefaults(e), false);
            });

            // Add visual feedback on drag enter/over
            ['dragenter', 'dragover'].forEach(eventName => {
                fileLabel.addEventListener(eventName, () => {
                    fileLabel.style.borderColor = '#3b82f6'; // Tailwind blue-500
                    fileLabel.style.backgroundColor = '#eff6ff'; // Tailwind blue-50
                }, false);
            });

            // Remove visual feedback on drag leave/drop
            ['dragleave', 'drop'].forEach(eventName => {
                fileLabel.addEventListener(eventName, () => {
                    fileLabel.style.borderColor = '#d1d5db'; // Tailwind gray-300
                    fileLabel.style.backgroundColor = '';
                }, false);
            });

            // Handle file drop
            fileLabel.addEventListener('drop', (e) => {
                const files = Array.from(e.dataTransfer.files).filter(file =>
                    file.name.toLowerCase().endsWith('.docx')
                );
                if (files.length > 0) {
                    // Create a DataTransfer object and assign files to the input
                    // This simulates a user selecting files via the input field
                    const dt = new DataTransfer();
                    files.forEach(file => dt.items.add(file));
                    wordFiles.files = dt.files;

                    // Manually dispatch a change event to trigger handleFileSelect
                    const event = new Event('change', { bubbles: true });
                    wordFiles.dispatchEvent(event);
                } else if (e.dataTransfer.files.length > 0) {
                    this.showNotification('Please upload only .docx files', 'warning');
                }
            }, false);
        }

        /**
         * Prevents default event behaviors (e.g., opening dropped files).
         * @param {Event} e - The event object.
         */
        preventDefaults(e) {
            e.preventDefault();
            e.stopPropagation();
        }

        /**
         * Handles file selection from the input or drag-and-drop.
         * Validates file types and updates the UI accordingly.
         * @param {Event} event - The change event from the file input.
         */
        handleFileSelect(event) {
            const files = Array.from(event.target.files);
            const fileLabel = document.getElementById('fileLabel');
            const processBtn = document.getElementById('processBtn');

            // Validate file types
            const validFiles = files.filter(file => file.name.toLowerCase().endsWith('.docx'));
            const invalidFiles = files.filter(file => !file.name.toLowerCase().endsWith('.docx'));

            if (invalidFiles.length > 0) {
                this.showNotification(`${invalidFiles.length} invalid file(s) ignored. Only .docx files are supported.`, 'warning');
            }

            if (validFiles.length > 0) {
                this.selectedFiles = validFiles;
                // Generate a concise list of file names for display
                const fileNames = validFiles.length > 3
                    ? `${validFiles.slice(0, 3).map(f => f.name).join(', ')} and ${validFiles.length - 3} more...`
                    : validFiles.map(f => f.name).join(', ');

                fileLabel.innerHTML = `
                    <span>📄</span>
                    <span>Selected: ${validFiles.length} document${validFiles.length > 1 ? 's' : ''}</span>
                    <small>${this.escapeHtml(fileNames)}</small>
                `;
                fileLabel.classList.add('has-files');
                processBtn.disabled = false;
                this.announceToScreenReader('process-status', `${validFiles.length} documents selected, ready to process`);
            } else {
                this.selectedFiles = [];
                fileLabel.innerHTML = `
                    <span>📄</span>
                    <span>Click here or drag Word documents to upload</span>
                    <small>Multiple files supported</small>
                `;
                fileLabel.classList.remove('has-files');
                processBtn.disabled = true;
                this.announceToScreenReader('process-status', 'No valid documents selected');
            }
        }

        /**
         * Processes the selected Word documents to extract poems.
         * Displays progress and notifications.
         */
        async processDocuments() {
            if (this.selectedFiles.length === 0) {
                this.showNotification('Please select Word documents first!', 'warning');
                return;
            }

            if (this.isProcessing) {
                return; // Prevent multiple simultaneous processing
            }

            this.isProcessing = true;
            const processBtn = document.getElementById('processBtn');
            const progressContainer = document.getElementById('progressContainer');
            const progressBar = document.getElementById('progressBar');

            // Update UI for processing state
            processBtn.disabled = true;
            processBtn.textContent = 'Processing...';
            progressContainer.style.display = 'block';
            progressBar.style.width = '0%';
            progressBar.setAttribute('aria-valuenow', '0');

            this.announceToScreenReader('process-status', 'Processing documents...');

            try {
                let processedPoemCount = 0;
                let skippedCount = 0;
                const totalFiles = this.selectedFiles.length;
                const errors = [];

                for (let i = 0; i < this.selectedFiles.length; i++) {
                    const file = this.selectedFiles[i];

                    try {
                        const poemsFromFile = await this.extractPoemsFromDocument(file);
                        if (poemsFromFile && poemsFromFile.length > 0) {
                            // Check for duplicates and add new poems
                            for (const poemData of poemsFromFile) {
                                if (poemData && poemData.content && poemData.content.trim().length > 0) {
                                    // Check for duplicates based on title and content similarity
                                    const isDuplicate = this.poems.some(existing =>
                                        existing.title.toLowerCase() === poemData.title.toLowerCase() &&
                                        (existing.content.trim().length > 50 && existing.content.trim() === poemData.content.trim())
                                    );

                                    if (!isDuplicate) {
                                        this.poems.push(poemData);
                                        processedPoemCount++;
                                    } else {
                                        skippedCount++;
                                        console.warn(`Duplicate poem detected and skipped: ${poemData.title || 'Untitled'}`);
                                    }
                                }
                            }
                        } else {
                            errors.push(`${file.name}: No valid poems found`);
                        }
                    } catch (error) {
                        console.error(`Error processing ${file.name}:`, error);
                        errors.push(`${file.name}: ${error.message}`);
                    }

                    // Update progress
                    const progress = ((i + 1) / totalFiles) * 100;
                    progressBar.style.width = `${progress}%`;
                    progressBar.setAttribute('aria-valuenow', Math.round(progress).toString());

                    // Small delay to show progress, but not block UI
                    await new Promise(resolve => requestAnimationFrame(resolve));
                }

                // Reset UI state
                this.resetProcessingUI();

                // Show results
                if (processedPoemCount > 0) {
                    this.updateDisplay();
                    let message = `Successfully processed ${processedPoemCount} new poem${processedPoemCount > 1 ? 's' : ''}!`;
                    if (skippedCount > 0) {
                        message += ` (${skippedCount} duplicate${skippedCount > 1 ? 's' : ''} skipped)`;
                    }
                    this.showNotification(message, 'success');
                    this.announceToScreenReader('process-status', `${processedPoemCount} poems processed successfully`);

                    // Reset file input only if some poems were successfully processed
                    this.resetFileInput();
                } else {
                    let message = 'No new poems found in the uploaded documents!';
                    if (skippedCount > 0) {
                        message = `All uploaded poems were duplicates or had no new content.`;
                    }
                    this.showNotification(message, 'warning');
                    this.announceToScreenReader('process-status', 'No new poems found');
                }

                // Show errors if any
                if (errors.length > 0) {
                    console.error('Processing errors:', errors);
                    this.showNotification(`${errors.length} file(s) had errors. Check console for details.`, 'error', 8000);
                }

            } catch (error) {
                this.resetProcessingUI();
                console.error('Document processing error:', error);
                this.showNotification('Error processing documents: ' + error.message, 'error');
                this.announceToScreenReader('process-status', 'Error processing documents');
            }
        }

        /**
         * Resets the UI elements related to document processing.
         */
        resetProcessingUI() {
            const processBtn = document.getElementById('processBtn');
            const progressContainer = document.getElementById('progressContainer');

            progressContainer.style.display = 'none';
            processBtn.textContent = 'Process Documents';
            processBtn.disabled = this.selectedFiles.length === 0; // Re-enable if files are selected
            this.isProcessing = false;
        }

        /**
         * Clears the selected files from the input and resets the file label.
         */
        resetFileInput() {
            const wordFiles = document.getElementById('wordFiles');
            if (wordFiles) {
                wordFiles.value = ''; // Clears the selected file(s) from the input
                // Manually trigger handleFileSelect with empty files to reset the label
                this.handleFileSelect({ target: { files: [] } });
            }
        }

        /**
         * Clears all loaded poems and updates the display.
         */
        clearAllPoems() {
            this.poems = [];
            this.updateDisplay();
            this.resetFileInput();
            this.showNotification('All poems cleared!', 'info');
            this.announceToScreenReader('process-status', 'All poems cleared.');
        }

        /**
         * Extracts HTML content from a DOCX file using Mammoth.js
         * and attempts to identify multiple poems within it.
         * @param {File} file - The DOCX file to process.
         * @returns {Promise<Array<Object>>} A promise resolving to an array of poem objects.
         * @throws {Error} If Mammoth.js is not loaded or content extraction fails.
         */
        async extractPoemsFromDocument(file) {
            if (!window.mammoth) {
                throw new Error('Mammoth library not loaded. Please check the script tag.');
            }

            try {
                const arrayBuffer = await file.arrayBuffer();
                const result = await window.mammoth.convertToHtml({ arrayBuffer });

                if (!result.value) {
                    throw new Error('No content extracted from document by Mammoth.');
                }

                const html = result.value;
                const tempDiv = document.createElement('div');
                tempDiv.innerHTML = html;

                // Get full content for analysis (plain text)
                const fullContent = tempDiv.textContent.trim();

                if (!fullContent || fullContent.length < 10) {
                    throw new Error('Document appears to be empty or too short after extraction.');
                }

                // Attempt to extract multiple poems from the document
                const poems = this.identifyMultiplePoems(tempDiv, file.name, html);

                if (poems.length === 0) {
                    // If no clear multiple poems were identified, treat the entire document as one poem
                    const singlePoem = this.createSinglePoemFromDocument(tempDiv, file.name, html, fullContent);
                    return [singlePoem];
                }

                return poems;

            } catch (error) {
                throw new Error(`Failed to extract content from "${file.name}": ${error.message}`);
            }
        }

        /**
         * Attempts to identify and separate multiple poems within an HTML document structure.
         * Uses different strategies (headings, paragraph breaks, separators).
         * @param {HTMLElement} tempDiv - A temporary div containing the document's HTML.
         * @param {string} filename - The original filename.
         * @param {string} fullHtml - The full HTML content from Mammoth.js.
         * @returns {Array<Object>} An array of identified poem objects.
         */
        identifyMultiplePoems(tempDiv, filename, fullHtml) {
            const poems = [];

            // Strategy 1: Split by headings (H1, H2, H3)
            const headings = tempDiv.querySelectorAll('h1, h2, h3');
            if (headings.length > 1) {
                const extractedPoems = this.extractPoemsByHeadings(tempDiv, filename, headings);
                if (extractedPoems.length > 1) return extractedPoems;
            }

            // Strategy 2: Split by multiple line breaks or page breaks (empty paragraphs)
            // This is done by looking at consecutive empty or very short paragraphs
            const paragraphs = Array.from(tempDiv.querySelectorAll('p'));
            if (paragraphs.length > 3) {
                const extractedPoems = this.extractPoemsByParagraphSeparation(tempDiv, filename, paragraphs);
                if (extractedPoems.length > 1) return extractedPoems;
            }

            // Strategy 3: Split by patterns like "***", "---", or similar visual separators
            const textContent = tempDiv.textContent;
            const separatorPatterns = [
                /\n\s*\*{3,}\s*\n/g,      // Three or more asterisks with surrounding newlines
                /\n\s*-{3,}\s*\n/g,      // Three or more dashes with surrounding newlines
                /\n\s*_{3,}\s*\n/g,      // Three or more underscores with surrounding newlines
                /\n\s*={3,}\s*\n/g,      // Three or more equals with surrounding newlines
                /\n\s*~{3,}\s*\n/g,      // Three or more tildes with surrounding newlines
                /\n\s*\n\s*\n\s*\n/g     // Four or more consecutive newlines (more explicit break)
            ];

            for (const pattern of separatorPatterns) {
                // Split the raw HTML to preserve formatting as much as possible
                const partsHtml = fullHtml.split(pattern);
                if (partsHtml.length > 1) {
                    const extractedPoems = this.extractPoemsBySeparator(partsHtml, filename);
                    if (extractedPoems.length > 1) return extractedPoems;
                }
            }

            return []; // Return empty array if no multiple poems detected by any strategy
        }

        /**
         * Extracts poems by identifying text blocks separated by heading tags (h1, h2, h3).
         * @param {HTMLElement} tempDiv - The temporary div containing the document HTML.
         * @param {string} filename - The name of the original file.
         * @param {NodeList<HTMLElement>} headings - A NodeList of h1, h2, h3 elements.
         * @returns {Array<Object>} An array of poem objects.
         */
        extractPoemsByHeadings(tempDiv, filename, headings) {
            const poems = [];
            const allElements = Array.from(tempDiv.children);

            for (let i = 0; i < headings.length; i++) {
                const currentHeading = headings[i];
                const nextHeading = headings[i + 1];

                const title = currentHeading.textContent.trim() || `Poem ${i + 1}`;

                // Find content between this heading and the next
                const startIndex = allElements.indexOf(currentHeading);
                const endIndex = nextHeading ? allElements.indexOf(nextHeading) : allElements.length;

                const poemElements = allElements.slice(startIndex + 1, endIndex);
                const poemContent = poemElements.map(el => el.textContent).join('\n').trim();
                const poemHtml = poemElements.map(el => el.outerHTML).join('\n');

                if (poemContent.length > 10) { // Ensure sufficient content
                    poems.push(this.createPoemObject(title, poemContent, poemHtml, filename));
                }
            }

            return poems.length > 1 ? poems : []; // Only return if we actually found multiple distinct poems
        }

        /**
         * Extracts poems by identifying blocks of paragraphs separated by empty or very short paragraphs.
         * Attempts to infer titles from the first paragraph of a new block if it fits title criteria.
         * @param {HTMLElement} tempDiv - The temporary div containing the document HTML.
         * @param {string} filename - The name of the original file.
         * @param {NodeList<HTMLElement>} paragraphs - A NodeList of paragraph elements.
         * @returns {Array<Object>} An array of poem objects.
         */
        extractPoemsByParagraphSeparation(tempDiv, filename, paragraphs) {
            const poems = [];
            let currentPoemElements = [];
            let currentTitle = '';
            let poemIndex = 1;

            for (let i = 0; i < paragraphs.length; i++) {
                const p = paragraphs[i];
                const text = p.textContent.trim();

                // Check if this might be a title (short, possibly bold/centered, first letter capitalized)
                const mightBeTitle = text.length > 0 && text.length < 100 &&
                    (p.querySelector('strong') || p.querySelector('b') ||
                     p.style.textAlign === 'center' || /^[A-Z][^.!?]*$/.test(text));

                // Check for poem separator (empty paragraph or very short paragraph acting as a break)
                const isEmptyOrBreak = text.length === 0 || (text.length < 10 && currentPoemElements.length > 0);

                if (isEmptyOrBreak) {
                    // End current poem if we have content
                    if (currentPoemElements.length > 0) {
                        const poemContent = currentPoemElements.map(el => el.textContent).join('\n').trim();
                        const poemHtml = currentPoemElements.map(el => el.outerHTML).join('\n');
                        const title = currentTitle || `Poem ${poemIndex}`;

                        if (poemContent.length > 10) { // Ensure content is substantial
                            poems.push(this.createPoemObject(title, poemContent, poemHtml, filename));
                            poemIndex++;
                        }
                        currentPoemElements = [];
                        currentTitle = '';
                    }
                } else if (mightBeTitle && currentPoemElements.length === 0) {
                    // This might be a title for a new poem, and no content has been added yet for it
                    currentTitle = text;
                    currentPoemElements.push(p); // Include the title paragraph in poem content
                } else {
                    // Regular poem content
                    currentPoemElements.push(p);
                }
            }

            // Handle the last poem
            if (currentPoemElements.length > 0) {
                const poemContent = currentPoemElements.map(el => el.textContent).join('\n').trim();
                const poemHtml = currentPoemElements.map(el => el.outerHTML).join('\n');
                const title = currentTitle || `Poem ${poemIndex}`;

                if (poemContent.length > 10) {
                    poems.push(this.createPoemObject(title, poemContent, poemHtml, filename));
                }
            }

            return poems.length > 1 ? poems : []; // Only return if we found multiple poems
        }

        /**
         * Extracts poems by splitting the HTML content based on detected separator patterns.
         * @param {Array<string>} htmlParts - Array of HTML strings separated by a pattern.
         * @param {string} filename - The name of the original file.
         * @returns {Array<Object>} An array of poem objects.
         */
        extractPoemsBySeparator(htmlParts, filename) {
            const poems = [];
            htmlParts.forEach((part, index) => {
                const tempDiv = document.createElement('div');
                tempDiv.innerHTML = part.trim();
                const content = tempDiv.textContent.trim();

                if (content.length > 10) { // Ensure sufficient content
                    // Try to extract a title from the first non-empty line of the part
                    const lines = content.split('\n').map(line => line.trim()).filter(line => line.length > 0);
                    const firstLine = lines[0] || '';
                    const title = (firstLine.length > 0 && firstLine.length < 100) ?
                        firstLine : `Poem ${index + 1}`;

                    poems.push(this.createPoemObject(title, content, part.trim(), filename));
                }
            });
            return poems;
        }

        /**
         * Creates a single poem object from an entire document when multiple poems are not detected.
         * @param {HTMLElement} tempDiv - The temporary div containing the document HTML.
         * @param {string} filename - The original filename.
         * @param {string} html - The full HTML content from Mammoth.js.
         * @param {string} content - The full plain text content of the document.
         * @returns {Object} A single poem object.
         */
        createSinglePoemFromDocument(tempDiv, filename, html, content) {
            const title = this.extractTitle(tempDiv, filename);
            const wordCount = content.split(/\s+/).filter(word => word.length > 0).length;

            return {
                id: Date.now() + Math.random(), // Unique ID for each poem
                title: title,
                content: content,
                htmlContent: html,
                filename: filename,
                wordCount: wordCount,
                dateAdded: new Date().toISOString()
            };
        }

        /**
         * Creates a poem object with all necessary properties.
         * @param {string} title - The title of the poem.
         * @param {string} content - The plain text content of the poem.
         * @param {string} htmlContent - The HTML content of the poem.
         * @param {string} filename - The original filename from which the poem was extracted.
         * @returns {Object} The poem object.
         */
        createPoemObject(title, content, htmlContent, filename) {
            const wordCount = content.split(/\s+/).filter(word => word.length > 0).length;

            return {
                id: Date.now() + Math.random(),
                title: title,
                content: content,
                htmlContent: htmlContent,
                filename: filename,
                wordCount: wordCount,
                dateAdded: new Date().toISOString()
            };
        }

        /**
         * Extracts a title from the document HTML, using various heuristics.
         * @param {HTMLElement} tempDiv - The temporary div containing the document's HTML.
         * @param {string} filename - The original filename.
         * @returns {string} The extracted or generated title.
         */
        extractTitle(tempDiv, filename) {
            let title = '';

            // 1. Try headings (H1, H2, H3)
            const headings = tempDiv.querySelectorAll('h1, h2, h3');
            for (let i = 0; i < headings.length; i++) {
                const hText = headings[i].textContent.trim();
                if (hText.length > 0 && hText.length < 150) { // Limit title length
                    title = hText;
                    break;
                }
            }

            // 2. Try bold or centered text near the beginning
            if (!title) {
                const paragraphs = tempDiv.querySelectorAll('p');
                for (let i = 0; i < Math.min(3, paragraphs.length); i++) { // Check first few paragraphs
                    const p = paragraphs[i];
                    const pText = p.textContent.trim();
                    if (pText.length > 0 && pText.length < 150) {
                        const isBold = p.querySelector('strong, b') !== null;
                        const isCentered = p.style.textAlign === 'center';

                        if (isBold || isCentered) {
                            title = pText;
                            break;
                        }
                    }
                }
            }

            // 3. If no clear heading or bold/centered, try first non-empty line
            if (!title) {
                const paragraphs = tempDiv.querySelectorAll('p');
                if (paragraphs.length > 0) {
                    const firstParagraphText = paragraphs[0].textContent.trim();
                    if (firstParagraphText.length > 0) {
                        const firstLine = firstParagraphText.split('\n')[0].trim();
                        if (firstLine.length > 0 && firstLine.length < 150) {
                            title = firstLine;
                        }
                    }
                }
            }

            // 4. Fallback to filename (cleaned up)
            if (!title) {
                title = filename.replace(/\.docx$/i, '').replace(/[_-]/g, ' ').trim();
            }

            // Final cleanup and default
            title = title.replace(/\s+/g, ' ').trim(); // Replace multiple spaces with single space
            if (title.length > 150) {
                title = title.substring(0, 147) + '...';
            }

            if (!title) {
                title = "Untitled Poem";
            }

            return title;
        }

        /**
         * Updates the display of loaded poems and their count.
         * Attaches drag-and-drop and button event listeners to each poem element.
         */
        updateDisplay() {
            const poemList = document.getElementById('poemList');
            const poemCountSpan = document.getElementById('poemCount');
            const downloadBtn = document.getElementById('downloadBtn');
            const clearBtn = document.getElementById('clearBtn');

            if (!poemList || !poemCountSpan || !downloadBtn || !clearBtn) {
                console.error('Required display elements not found for updateDisplay');
                return;
            }

            poemList.innerHTML = ''; // Clear existing list items
            poemCountSpan.textContent = this.poems.length;

            if (this.poems.length === 0) {
                poemList.innerHTML = `
                    <div class="widget-empty-state">
                        <p>No poems loaded yet. Upload and process Word documents to begin.</p>
                    </div>
                `;
                downloadBtn.disabled = true;
                clearBtn.disabled = true;
            } else {
                this.poems.forEach((poem, index) => {
                    const poemDiv = this.createPoemElement(poem, index);
                    poemList.appendChild(poemDiv);
                });

                downloadBtn.disabled = false;
                clearBtn.disabled = false;
            }
        }

        /**
         * Creates a DOM element for a single poem to be displayed in the list.
         * @param {Object} poem - The poem object.
         * @param {number} index - The current index of the poem in the array.
         * @returns {HTMLElement} The created poem div element.
         */
        createPoemElement(poem, index) {
            const poemDiv = document.createElement('div');
            poemDiv.classList.add('widget-poem-item');
            poemDiv.setAttribute('draggable', 'true'); // Enable drag
            poemDiv.setAttribute('data-index', index); // Store index for reordering
            poemDiv.setAttribute('role', 'listitem');
            poemDiv.setAttribute('aria-label', `Poem: ${poem.title}, position ${index + 1} of ${this.poems.length}. Press Ctrl+Up/Down to move, Delete to remove.`);
            poemDiv.setAttribute('tabindex', '0'); // Make draggable items keyboard focusable

            // Create a safe preview text (truncated and HTML escaped)
            const preview = poem.content.length > 100
                ? poem.content.substring(0, 100).split('\n')[0] + '...' // Take first line of preview
                : poem.content.split('\n')[0]; // Take first line if short

            poemDiv.innerHTML = `
                <div class="widget-drag-indicator" aria-hidden="true">⋮⋮</div>
                <div class="widget-poem-details">
                    <h3>${this.escapeHtml(poem.title)}</h3>
                    <p><strong>Source:</strong> ${this.escapeHtml(poem.filename)}</p>
                    <p><strong>Word Count:</strong> ${poem.wordCount}</p>
                    <p><strong>Preview:</strong> ${this.escapeHtml(preview)}</p>
                </div>
                <div class="widget-poem-controls">
                    ${index > 0 ? `<button class="widget-move-btn"
                                data-index="${index}"
                                aria-label="Move ${this.escapeHtml(poem.title)} up in the list">
                            <span aria-hidden="true">↑</span>
                        </button>` : '<div style="width: 32px; visibility: hidden;"></div>'}
                    <button class="widget-remove-btn"
                            data-index="${index}"
                            aria-label="Remove ${this.escapeHtml(poem.title)} from the list">
                        <span aria-hidden="true">×</span>
                    </button>
                    ${index < this.poems.length - 1 ? `<button class="widget-move-btn move-down"
                                data-index="${index}"
                                aria-label="Move ${this.escapeHtml(poem.title)} down in the list">
                            <span aria-hidden="true">↓</span>
                        </button>` : '<div style="width: 32px; visibility: hidden;"></div>'}
                </div>
            `;

            this.attachPoemEventListeners(poemDiv, index);
            return poemDiv;
        }

        /**
         * Attaches drag-and-drop, move, and remove event listeners to a poem element.
         * @param {HTMLElement} poemDiv - The poem's DOM element.
         * @param {number} index - The current index of the poem.
         */
        attachPoemEventListeners(poemDiv, index) {
            // --- Drag and drop event listeners ---
            poemDiv.addEventListener('dragstart', (e) => {
                this.draggedIndex = index;
                poemDiv.classList.add('dragging');
                e.dataTransfer.effectAllowed = 'move';
                e.dataTransfer.setData('text/plain', index.toString()); // Store original index
                this.announceToScreenReader('process-status', `Started dragging ${this.poems[index].title}`);
            });

            poemDiv.addEventListener('dragend', () => {
                // Remove dragging class from all items
                document.querySelectorAll('.widget-poem-item').forEach(item => {
                    item.classList.remove('dragging');
                    item.classList.remove('drag-over');
                });
                this.draggedIndex = null;
            });

            poemDiv.addEventListener('dragover', (e) => {
                e.preventDefault(); // Allow drop
                e.dataTransfer.dropEffect = 'move';
                const targetElement = e.currentTarget;

                // Add drag-over class to visual feedback
                document.querySelectorAll('.widget-poem-item').forEach(item => {
                    item.classList.remove('drag-over');
                });
                if (targetElement.classList.contains('widget-poem-item') && this.draggedIndex !== null) {
                    const targetIndex = parseInt(targetElement.dataset.index);
                    if (targetIndex !== this.draggedIndex) {
                        targetElement.classList.add('drag-over');
                    }
                }
            });

            poemDiv.addEventListener('dragleave', (e) => {
                e.currentTarget.classList.remove('drag-over');
            });

            poemDiv.addEventListener('drop', (e) => {
                e.preventDefault();
                e.currentTarget.classList.remove('drag-over');
                const draggedIdx = parseInt(e.dataTransfer.getData('text/plain'));
                const dropTargetIndex = parseInt(e.currentTarget.dataset.index);

                if (draggedIdx !== dropTargetIndex && !isNaN(draggedIdx)) {
                    this.movePoem(draggedIdx, dropTargetIndex);
                }
            });

            // --- Keyboard accessibility for drag and drop (Ctrl + Arrow keys) ---
            poemDiv.addEventListener('keydown', (e) => {
                if (e.key === 'ArrowUp' && e.ctrlKey && index > 0) {
                    e.preventDefault(); // Prevent page scrolling
                    this.movePoem(index, index - 1);
                    // Re-focus the moved item to maintain accessibility
                    requestAnimationFrame(() => {
                        const newPoemDiv = document.querySelector(`.widget-poem-item[data-index="${index - 1}"]`);
                        if (newPoemDiv) newPoemDiv.focus();
                        this.announceToScreenReader('process-status', `Moved ${this.poems[index - 1].title} to position ${index}.`);
                    });
                } else if (e.key === 'ArrowDown' && e.ctrlKey && index < this.poems.length - 1) {
                    e.preventDefault(); // Prevent page scrolling
                    this.movePoem(index, index + 1);
                    // Re-focus the moved item
                    requestAnimationFrame(() => {
                        const newPoemDiv = document.querySelector(`.widget-poem-item[data-index="${index + 1}"]`);
                        if (newPoemDiv) newPoemDiv.focus();
                        this.announceToScreenReader('process-status', `Moved ${this.poems[index + 1].title} to position ${index + 2}.`);
                    });
                } else if (e.key === 'Delete' || e.key === 'Backspace') {
                    e.preventDefault(); // Prevent browser back navigation
                    const confirmed = true; // No confirm() as per guidelines. Can implement custom modal if needed.
                    if (confirmed) {
                        this.removePoem(index);
                        this.announceToScreenReader('process-status', `Removed ${this.poems[index].title}.`);
                    }
                }
            });

            // --- Click event listeners for move/remove buttons ---
            const moveUpBtn = poemDiv.querySelector('.widget-move-btn:not(.move-down)');
            if (moveUpBtn) {
                moveUpBtn.addEventListener('click', (e) => {
                    e.preventDefault();
                    if (index > 0) {
                        this.movePoem(index, index - 1);
                    }
                });
            }

            const moveDownBtn = poemDiv.querySelector('.widget-move-btn.move-down');
            if (moveDownBtn) {
                moveDownBtn.addEventListener('click', (e) => {
                    e.preventDefault();
                    if (index < this.poems.length - 1) {
                        this.movePoem(index, index + 1);
                    }
                });
            }

            const removeBtn = poemDiv.querySelector('.widget-remove-btn');
            if (removeBtn) {
                removeBtn.addEventListener('click', (e) => {
                    e.preventDefault();
                    const confirmed = true; // No confirm()
                    if (confirmed) {
                        this.removePoem(index);
                    }
                });
            }
        }

        /**
         * Safely escapes HTML special characters in a string to prevent XSS.
         * @param {string} text - The text to escape.
         * @returns {string} The HTML-escaped string.
         */
        escapeHtml(text) {
            if (typeof text !== 'string') return '';
            const div = document.createElement('div');
            div.textContent = text;
            return div.innerHTML;
        }

        /**
         * Moves a poem from one position to another in the array and updates the display.
         * @param {number} fromIndex - The original index of the poem.
         * @param {number} toIndex - The target index for the poem.
         */
        movePoem(fromIndex, toIndex) {
            if (fromIndex < 0 || fromIndex >= this.poems.length ||
                toIndex < 0 || toIndex >= this.poems.length) {
                console.error('Invalid indices for movePoem', fromIndex, toIndex);
                return;
            }

            const [movedPoem] = this.poems.splice(fromIndex, 1); // Remove from old position
            this.poems.splice(toIndex, 0, movedPoem); // Insert into new position
            this.updateDisplay(); // Re-render the list
            this.showNotification(`Moved "${movedPoem.title}" from position ${fromIndex + 1} to ${toIndex + 1}`, 'info');
            this.announceToScreenReader('process-status', `Poem moved. New order updated.`);
        }

        /**
         * Removes a poem from the array and updates the display.
         * @param {number} index - The index of the poem to remove.
         */
        removePoem(index) {
            if (index < 0 || index >= this.poems.length) {
                console.error('Invalid index for removePoem', index);
                return;
            }
            const removedPoem = this.poems.splice(index, 1)[0]; // Remove the poem
            this.updateDisplay(); // Re-render the list
            this.showNotification(`Removed "${removedPoem.title}"`, 'info');
            this.announceToScreenReader('process-status', `Poem ${removedPoem.title} removed.`);
        }

        /**
         * Displays a notification message to the user.
         * @param {string} message - The message to display.
         * @param {string} type - The type of notification (success, warning, error, info).
         * @param {number} [duration=5000] - Duration in milliseconds before the notification fades.
         */
        showNotification(message, type, duration = 5000) {
            const notificationContainer = document.getElementById('notificationContainer');
            if (!notificationContainer) return;

            // Clear any existing timeout
            if (this.notificationTimeout) {
                clearTimeout(this.notificationTimeout);
            }
            notificationContainer.innerHTML = ''; // Clear previous notifications

            const notificationDiv = document.createElement('div');
            notificationDiv.classList.add('widget-notification', type, 'opacity-0', 'transition-opacity', 'duration-300');
            notificationDiv.setAttribute('role', type === 'error' ? 'alert' : 'status');
            notificationDiv.innerHTML = `
                <span class="mr-2">${this._getNotificationIcon(type)}</span>
                <span>${this.escapeHtml(message)}</span>
            `;
            notificationContainer.appendChild(notificationDiv);

            // Trigger fade-in
            setTimeout(() => {
                notificationDiv.classList.remove('opacity-0');
            }, 10); // Small delay to allow render before transition

            // Set timeout to fade out and remove
            this.notificationTimeout = setTimeout(() => {
                notificationDiv.classList.add('opacity-0');
                notificationDiv.addEventListener('transitionend', () => {
                    if (notificationDiv.parentNode) {
                        notificationDiv.parentNode.removeChild(notificationDiv);
                    }
                }, { once: true });
            }, duration);
        }

        /**
         * Returns an icon based on notification type.
         * @param {string} type - The notification type.
         * @returns {string} An emoji icon.
         */
        _getNotificationIcon(type) {
            switch (type) {
                case 'success': return '✅';
                case 'warning': return '⚠️';
                case 'error': return '❌';
                case 'info': return 'ℹ️';
                default: return '';
            }
        }

        /**
         * Announces messages to screen readers for accessibility.
         * @param {string} elementId - The ID of the ARIA live region element.
         * @param {string} message - The message to announce.
         */
        announceToScreenReader(elementId, message) {
            const el = document.getElementById(elementId);
            if (el) {
                el.textContent = message;
            }
        }

        /**
         * Generates the table of contents HTML based on the current poem order.
         * @returns {string} The HTML string for the table of contents.
         */
        generateTableOfContentsHtml() {
            if (this.poems.length === 0) {
                return '';
            }

            let tocHtml = `<h2 style="text-align: center; margin-bottom: 20px; font-size: 2em; color: #333;">Table of Contents</h2>\n`;
            tocHtml += `<ol style="list-style-type: decimal; margin-left: 20px; line-height: 1.8;">\n`;
            this.poems.forEach((poem, index) => {
                // Ensure unique IDs for anchors
                const poemAnchorId = `poem-${index + 1}-${poem.id}`;
                tocHtml += `<li><a href="#${poemAnchorId}" style="color: #007bff; text-decoration: none;">${this.escapeHtml(poem.title)}</a></li>\n`;
            });
            tocHtml += `</ol>\n\n`;
            tocHtml += `<div style="page-break-after: always;"></div>\n`; // Page break after TOC
            return tocHtml;
        }

        /**
         * Downloads the combined document in the selected format.
         */
        async downloadCombinedDocument() {
            if (this.poems.length === 0) {
                this.showNotification('No poems to download!', 'warning');
                return;
            }

            const exportFormat = document.getElementById('exportFormat').value;
            const downloadBtn = document.getElementById('downloadBtn');

            // Set UI to downloading state
            downloadBtn.disabled = true;
            const originalText = downloadBtn.textContent;
            downloadBtn.textContent = 'Generating...';
            this.showNotification(`Generating ${exportFormat.toUpperCase()}...`, 'info', 0); // Indefinite until done

            try {
                let filename = 'Combined_Poems';
                let blob;

                const combinedHtml = this._generateHtmlContentForExport();

                switch (exportFormat) {
                    case 'html':
                        blob = new Blob([combinedHtml], { type: 'text/html' });
                        filename += '.html';
                        break;
                    case 'docx':
                        // For .docx, we're essentially saving an HTML file with a .doc/.docx extension.
                        // Word is generally good at opening these.
                        blob = new Blob([combinedHtml], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
                        filename += '.docx';
                        this.showNotification('Downloading .docx. Please note: This is an HTML file saved as .docx for compatibility with Word. Formatting may vary.', 'info', 10000);
                        break;
                    case 'pdf':
                        // Use html2pdf.js for PDF generation
                        await this._generatePdfOutput(combinedHtml, filename);
                        this.showNotification('PDF generated!', 'success');
                        this.resetDownloadUI(downloadBtn, originalText);
                        return; // Exit as html2pdf handles saving
                    default:
                        this.showNotification('Invalid export format selected.', 'error');
                        return;
                }

                saveAs(blob, filename); // Uses FileSaver.js
                this.showNotification('Document downloaded successfully!', 'success');

            } catch (error) {
                console.error('Error during document download:', error);
                this.showNotification('Error downloading document: ' + error.message, 'error');
            } finally {
                this.resetDownloadUI(downloadBtn, originalText);
            }
        }

        /**
         * Resets the download button UI after generation.
         * @param {HTMLElement} downloadBtn - The download button element.
         * @param {string} originalText - The original text content of the button.
         */
        resetDownloadUI(downloadBtn, originalText) {
            downloadBtn.textContent = originalText;
            downloadBtn.disabled = this.poems.length === 0;
            // Clear the indefinite notification
            const notificationContainer = document.getElementById('notificationContainer');
            if (notificationContainer) {
                notificationContainer.innerHTML = '';
            }
        }

        /**
         * Generates the full HTML content including TOC and poems with styling.
         * @returns {string} The complete HTML string.
         */
        _generateHtmlContentForExport() {
            let combinedHtml = `
            <!DOCTYPE html>
            <html lang="en">
            <head>
                <meta charset="UTF-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <title>Combined Poems</title>
                <style>
                    body {
                        font-family: 'Inter', sans-serif;
                        line-height: 1.6;
                        margin: 20px auto;
                        max-width: 800px;
                        padding: 0 20px;
                        color: #333;
                    }
                    h1 {
                        text-align: center;
                        font-size: 3em;
                        color: #1a202c;
                        margin-bottom: 30px;
                    }
                    h2 {
                        font-size: 2em;
                        color: #333;
                        margin-top: 40px;
                        margin-bottom: 15px;
                        border-bottom: 1px solid #eee;
                        padding-bottom: 5px;
                    }
                    h3 {
                        font-size: 1.5em;
                        color: #444;
                        margin-top: 30px;
                        margin-bottom: 10px;
                    }
                    p {
                        margin-bottom: 1em;
                    }
                    .poem-container {
                        margin-bottom: 40px;
                        padding-bottom: 20px;
                        border-bottom: 1px dashed #ddd;
                        page-break-inside: avoid; /* Prevent page breaks within a poem */
                    }
                    .poem-container:last-of-type {
                        border-bottom: none;
                        margin-bottom: 0;
                    }
                    .poem-source {
                        font-style: italic;
                        color: #666;
                        font-size: 0.9em;
                        margin-top: -10px;
                        margin-bottom: 15px;
                    }
                    .table-of-contents {
                        margin-bottom: 50px;
                        padding: 20px;
                        background-color: #f9f9f9;
                        border: 1px solid #eee;
                        border-radius: 8px;
                    }
                    .table-of-contents ol {
                        list-style-type: decimal;
                        padding-left: 25px;
                    }
                    .table-of-contents li {
                        margin-bottom: 5px;
                    }
                    .table-of-contents a {
                        color: #007bff;
                        text-decoration: none;
                        transition: color 0.2s ease-in-out;
                    }
                    .table-of-contents a:hover {
                        color: #0056b3;
                        text-decoration: underline;
                    }
                    /* Page break for printing/PDF */
                    .page-break-after {
                        page-break-after: always;
                    }
                </style>
            </head>
            <body>
                <h1>A Collection of Poems</h1>
                <div class="table-of-contents">
                    ${this.generateTableOfContentsHtml()}
                </div>
            `;

            this.poems.forEach((poem, index) => {
                const poemAnchorId = `poem-${index + 1}-${poem.id}`; // Match TOC anchor IDs
                combinedHtml += `
                <div class="poem-container" id="${poemAnchorId}">
                    <h2>${this.escapeHtml(poem.title)}</h2>
                    <p class="poem-source">From: ${this.escapeHtml(poem.filename)}</p>
                    ${poem.htmlContent}
                </div>
                `;
                // Add page break after each poem for better printing/PDF layout, unless it's the last one
                if (index < this.poems.length - 1) {
                    combinedHtml += `<div class="page-break-after"></div>`;
                }
            });

            combinedHtml += `
            </body>
            </html>
            `;
            return combinedHtml;
        }

        /**
         * Generates and downloads a PDF document from the given HTML content.
         * @param {string} htmlContent - The HTML string to convert to PDF.
         * @param {string} filename - The desired filename for the PDF.
         */
        async _generatePdfOutput(htmlContent, filename) {
            const opt = {
                margin: [20, 20, 20, 20], // Top, Left, Bottom, Right
                filename: filename,
                image: { type: 'jpeg', quality: 0.98 },
                html2canvas: { scale: 2 },
                jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
            };

            // Create a temporary element to render the HTML for html2pdf
            const tempDiv = document.createElement('div');
            tempDiv.innerHTML = htmlContent;
            tempDiv.style.width = '210mm'; // Set a fixed width for A4
            tempDiv.style.margin = '0 auto';
            tempDiv.style.visibility = 'hidden'; // Hide it from user view
            document.body.appendChild(tempDiv);

            try {
                // Use html2pdf to convert the tempDiv content
                await html2pdf().set(opt).from(tempDiv).save();
            } finally {
                // Clean up the temporary element
                if (tempDiv.parentNode) {
                    document.body.removeChild(tempDiv);
                }
            }
        }
    }

    // Initialize the PoemCompiler once the DOM is fully loaded
    document.addEventListener('DOMContentLoaded', () => {
        new PoemCompiler();
    });
})();
