(function() {
    'use strict';

    class PoemCompiler {
        constructor() {
            this.poems = [];
            this.selectedFiles = [];
            this.draggedIndex = null;
            this.isProcessing = false;
            this.notificationTimeout = null;
            console.log('PoemCompiler initialized.');
            this.initializeEventListeners();
            this.updateDisplay();
        }

        /**
         * Initializes all event listeners for the UI elements.
         */
        initializeEventListeners() {
            console.log('Initializing event listeners...');
            const wordFiles = document.getElementById('wordFiles');
            const processBtn = document.getElementById('processBtn');
            const downloadBtn = document.getElementById('downloadBtn');
            const clearBtn = document.getElementById('clearBtn');
            const fileLabel = document.getElementById('fileLabel');

            if (!wordFiles || !processBtn || !downloadBtn || !clearBtn || !fileLabel) {
                console.error('Required DOM elements not found. Please ensure all IDs are correct in the HTML.');
                this.showNotification('Application setup error: Some UI elements are missing. Please check the HTML.', 'error', 0);
                return;
            }

            // File input change event
            wordFiles.addEventListener('change', (e) => {
                console.log('File input change event triggered.');
                this.handleFileSelect(e);
            });

            // Process button click event
            processBtn.addEventListener('click', () => {
                console.log('Process button clicked. current this.isProcessing:', this.isProcessing);
                if (!this.isProcessing) {
                    this.processDocuments();
                } else {
                    this.showNotification('Already processing documents. Please wait.', 'info');
                }
            });

            // Download button click event
            downloadBtn.addEventListener('click', () => {
                console.log('Download button clicked.');
                this.downloadCombinedDocument();
            });

            // Clear button click event
            clearBtn.addEventListener('click', () => {
                console.log('Clear button clicked.');
                this.clearAllPoems();
            });

            // --- Drag and drop functionality for the file label ---
            ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
                fileLabel.addEventListener(eventName, (e) => this.preventDefaults(e), false);
            });

            ['dragenter', 'dragover'].forEach(eventName => {
                fileLabel.addEventListener(eventName, () => {
                    fileLabel.style.borderColor = '#3b82f6';
                    fileLabel.style.backgroundColor = '#eff6ff';
                }, false);
            });

            ['dragleave', 'drop'].forEach(eventName => {
                fileLabel.addEventListener(eventName, () => {
                    fileLabel.style.borderColor = '#d1d5db';
                    fileLabel.style.backgroundColor = '';
                }, false);
            });

            fileLabel.addEventListener('drop', (e) => {
                console.log('File dropped.');
                const files = Array.from(e.dataTransfer.files).filter(file =>
                    file.name.toLowerCase().endsWith('.docx')
                );
                if (files.length > 0) {
                    const dt = new DataTransfer();
                    files.forEach(file => dt.items.add(file));
                    wordFiles.files = dt.files;
                    const event = new Event('change', { bubbles: true });
                    wordFiles.dispatchEvent(event); // Trigger change event programmatically
                    console.log('Dropped .docx files, dispatched change event.');
                } else if (e.dataTransfer.files.length > 0) {
                    this.showNotification('Please upload only .docx files', 'warning');
                    console.warn('Dropped files but none were .docx.');
                }
            }, false);
            console.log('Event listeners initialized successfully.');
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
            console.log('Files selected in handleFileSelect:', files);

            const fileLabel = document.getElementById('fileLabel');
            const processBtn = document.getElementById('processBtn');

            const validFiles = files.filter(file => file.name.toLowerCase().endsWith('.docx'));
            const invalidFiles = files.filter(file => !file.name.toLowerCase().endsWith('.docx'));

            if (invalidFiles.length > 0) {
                this.showNotification(`${invalidFiles.length} invalid file(s) ignored. Only .docx files are supported.`, 'warning');
                console.warn(`${invalidFiles.length} invalid file(s) ignored.`);
            }

            if (validFiles.length > 0) {
                this.selectedFiles = validFiles;
                console.log('Valid files assigned to this.selectedFiles:', this.selectedFiles);
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
                console.log(`UI updated: ${validFiles.length} documents selected, process button enabled.`);
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
                console.log('No valid files selected. UI reset.');
            }
        }

        /**
         * Processes the selected Word documents to extract poems.
         * Displays progress and notifications.
         */
        async processDocuments() {
            if (this.selectedFiles.length === 0) {
                this.showNotification('Please select Word documents first!', 'warning');
                console.warn('ProcessDocuments called with no selected files.');
                return;
            }

            if (this.isProcessing) {
                console.log('Attempted to process while already processing.');
                return;
            }

            this.isProcessing = true;
            const processBtn = document.getElementById('processBtn');
            const progressContainer = document.getElementById('progressContainer');
            const progressBar = document.getElementById('progressBar');

            console.log('Starting processDocuments. Updating UI...');
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
                console.log(`Processing ${totalFiles} selected files.`);
                const errors = [];

                for (let i = 0; i < this.selectedFiles.length; i++) {
                    const file = this.selectedFiles[i];
                    console.log(`Processing file ${i + 1}/${totalFiles}: ${file.name}`);

                    try {
                        const poemsFromFile = await this.extractPoemsFromDocument(file);
                        console.log(`Extracted ${poemsFromFile ? poemsFromFile.length : 0} potential poems from ${file.name}`);
                        if (poemsFromFile && poemsFromFile.length > 0) {
                            for (const poemData of poemsFromFile) {
                                if (poemData && poemData.content && poemData.content.trim().length > 0) {
                                    const isDuplicate = this.poems.some(existing =>
                                        existing.title.toLowerCase() === poemData.title.toLowerCase() &&
                                        (existing.content.trim().length > 50 && existing.content.trim() === poemData.content.trim())
                                    );

                                    if (!isDuplicate) {
                                        this.poems.push(poemData);
                                        processedPoemCount++;
                                        console.log(`Added new poem: "${poemData.title}" from "${file.name}"`);
                                    } else {
                                        skippedCount++;
                                        console.warn(`Duplicate poem detected and skipped: "${poemData.title || 'Untitled'}" from "${file.name}"`);
                                    }
                                } else {
                                    console.warn(`Poem data from ${file.name} was empty or invalid.`);
                                }
                            }
                        } else {
                            errors.push(`${file.name}: No valid poems found`);
                            console.warn(`No valid poems found in ${file.name}.`);
                        }
                    } catch (error) {
                        console.error(`Error processing ${file.name}:`, error);
                        errors.push(`${file.name}: ${error.message}`);
                    }

                    const progress = ((i + 1) / totalFiles) * 100;
                    progressBar.style.width = `${progress}%`;
                    progressBar.setAttribute('aria-valuenow', Math.round(progress).toString());
                    console.log(`Progress: ${Math.round(progress)}%`);

                    await new Promise(resolve => requestAnimationFrame(resolve));
                }

                console.log('Finished processing all files. Resetting UI.');
                this.resetProcessingUI();

                if (processedPoemCount > 0) {
                    this.updateDisplay();
                    let message = `Successfully processed ${processedPoemCount} new poem${processedPoemCount > 1 ? 's' : ''}!`;
                    if (skippedCount > 0) {
                        message += ` (${skippedCount} duplicate${skippedCount > 1 ? 's' : ''} skipped)`;
                    }
                    this.showNotification(message, 'success');
                    this.announceToScreenReader('process-status', `${processedPoemCount} poems processed successfully`);
                    this.resetFileInput();
                    console.log('Poem processing complete. Display updated.');
                } else {
                    let message = 'No new poems found in the uploaded documents!';
                    if (skippedCount > 0) {
                        message = `All uploaded poems were duplicates or had no new content.`;
                    }
                    this.showNotification(message, 'warning');
                    this.announceToScreenReader('process-status', 'No new poems found');
                    console.log('No new poems added after processing.');
                }

                if (errors.length > 0) {
                    console.error('Summary of processing errors:', errors);
                    this.showNotification(`${errors.length} file(s) had errors. Check console for details.`, 'error', 8000);
                }

            } catch (error) {
                this.resetProcessingUI();
                console.error('Unhandled critical error during document processing:', error);
                this.showNotification('A critical error occurred: ' + error.message, 'error');
                this.announceToScreenReader('process-status', 'Critical error during document processing.');
            }
        }

        /**
         * Resets the UI elements related to document processing.
         */
        resetProcessingUI() {
            console.log('Resetting processing UI.');
            const processBtn = document.getElementById('processBtn');
            const progressContainer = document.getElementById('progressContainer');

            progressContainer.style.display = 'none';
            processBtn.textContent = 'Process Documents';
            processBtn.disabled = this.selectedFiles.length === 0;
            this.isProcessing = false;
        }

        /**
         * Clears the selected files from the input and resets the file label.
         */
        resetFileInput() {
            console.log('Resetting file input.');
            const wordFiles = document.getElementById('wordFiles');
            if (wordFiles) {
                wordFiles.value = '';
                this.handleFileSelect({ target: { files: [] } });
            }
        }

        /**
         * Clears all loaded poems and updates the display.
         */
        clearAllPoems() {
            console.log('Clearing all poems.');
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
            console.log(`Attempting to extract poems from "${file.name}"...`);
            if (!window.mammoth) {
                console.error('Mammoth library (window.mammoth) is not loaded.');
                throw new Error('Mammoth library not loaded. Please check the script tag.');
            }

            try {
                const arrayBuffer = await file.arrayBuffer();
                console.log(`File "${file.name}" converted to ArrayBuffer.`);
                const result = await window.mammoth.convertToHtml({ arrayBuffer });
                console.log(`Mammoth conversion result for "${file.name}":`, result);

                if (!result.value) {
                    console.warn(`Mammoth returned no HTML content for "${file.name}".`);
                    throw new Error('No content extracted from document by Mammoth.');
                }

                const html = result.value;
                const tempDiv = document.createElement('div');
                tempDiv.innerHTML = html;
                const fullContent = tempDiv.textContent.trim();
                const preservedHtml = this.preserveFormattingInHtml(html); // MANUAL FIX 1
                console.log(`Full plain text content length for "${file.name}": ${fullContent.length}`);

                if (!fullContent || fullContent.length < 10) {
                    console.warn(`Document "${file.name}" appears empty or too short after extraction.`);
                    throw new Error('Document appears to be empty or too short after extraction.');
                }

                const poems = this.identifyMultiplePoems(tempDiv, file.name, html);
                console.log(`identifyMultiplePoems returned ${poems.length} poems for "${file.name}".`);

                if (poems.length === 0) {
                    const singlePoem = this.createSinglePoemFromDocument(tempDiv, file.name, preservedHtml, fullContent); // Use preservedHtml
                    console.log(`No multiple poems detected, treating "${file.name}" as a single poem: "${singlePoem.title}".`);
                    return [singlePoem];
                }

                return poems;

            } catch (error) {
                console.error(`Failed to extract content from "${file.name}":`, error);
                throw new Error(`Failed to extract content from "${file.name}": ${error.message}`);
            }
        }

        /**
         * Preserves critical formatting elements from Mammoth HTML
         * @param {string} html - The HTML content from Mammoth
         * @returns {string} HTML with preserved formatting
         */
        preserveFormattingInHtml(html) { // MANUAL FIX 2
            // Replace multiple consecutive spaces with non-breaking spaces
            let formatted = html.replace(/  +/g, (match) => '&nbsp;'.repeat(match.length));

            // Preserve indentation by converting leading spaces to non-breaking spaces
            formatted = formatted.replace(/^( +)/gm, (match) => '&nbsp;'.repeat(match.length));

            // Ensure line breaks are preserved
            formatted = formatted.replace(/\n/g, '<br>');

            // Preserve paragraph spacing
            formatted = formatted.replace(/<\/p>\s*<p>/g, '</p><p style="margin-top: 1em;">');

            return formatted;
        }

        /**
         * Improves the poem identification to avoid fragmentation.
         * @param {HTMLElement} tempDiv - A temporary div containing the document's HTML.
         * @param {string} filename - The original filename.
         * @param {string} fullHtml - The full HTML content from Mammoth.js.
         * @returns {Array<Object>} An array of identified poem objects.
         */
        identifyMultiplePoems(tempDiv, filename, fullHtml) { // MANUAL FIX 3
            console.log(`Starting identifyMultiplePoems for "${filename}".`);

            // First, try to detect if this is a collection vs single poem
            const textContent = tempDiv.textContent;
            const lineCount = textContent.split('\n').filter(line => line.trim().length > 0).length;

            // If document is relatively short (under 50 meaningful lines), treat as single poem
            if (lineCount < 50) {
                console.log(`Document appears to be a single poem (${lineCount} lines)`);
                return [];
            }

            // Strategy 1: Split by clear title patterns (bold, centered, or all caps headers)
            const headings = tempDiv.querySelectorAll('h1, h2, h3, p strong, p b');
            if (headings.length > 1) {
                const extractedPoems = this.extractPoemsByStrongTitles(tempDiv, filename, headings);
                if (extractedPoems.length > 1) {
                    console.log(`Strategy 1 (Strong Titles) found ${extractedPoems.length} poems.`);
                    return extractedPoems;
                }
            }

            // Strategy 2: Split by significant whitespace gaps (3+ empty lines)
            const significantBreaks = fullHtml.split(/(<p[^>]*>\s*<\/p>\s*){3,}/);
            if (significantBreaks.length > 2) {
                const extractedPoems = this.extractPoemsBySignificantBreaks(significantBreaks, filename);
                if (extractedPoems.length > 1) {
                    console.log(`Strategy 2 (Significant Breaks) found ${extractedPoems.length} poems.`);
                    return extractedPoems;
                }
            }

            // If no clear separation found, treat as single document
            console.log(`No clear poem separation found for "${filename}".`);
            return [];
        }

        /**
         * Extracts poems based on strong visual titles (bold, headers)
         */
        extractPoemsByStrongTitles(tempDiv, filename, titleElements) { // MANUAL FIX 4
            const poems = [];
            const allElements = Array.from(tempDiv.children);

            for (let i = 0; i < titleElements.length; i++) {
                const currentTitle = titleElements[i];
                const nextTitle = titleElements[i + 1];

                const titleText = currentTitle.textContent.trim();
                if (!titleText || titleText.length > 100) continue;

                // Determine the starting element for the poem content
                // If the title element is a heading (h1-h3) or directly within a p tag,
                // find its closest parent paragraph for starting slice.
                const startIndex = allElements.indexOf(currentTitle.closest('p') || currentTitle);
                const endIndex = nextTitle ?
                    allElements.indexOf(nextTitle.closest('p') || nextTitle) :
                    allElements.length;

                const poemElements = allElements.slice(startIndex, endIndex); // Include the title element itself
                const poemHtml = poemElements.map(el => el.outerHTML).join('\n');
                const poemContent = poemElements.map(el => el.textContent).join('\n').trim();

                if (poemContent.length > 20) { // Ensure sufficient content to be a poem
                    poems.push(this.createPoemObject(titleText, poemContent, this.preserveFormattingInHtml(poemHtml), filename));
                }
            }

            return poems.length > 1 ? poems : [];
        }

        /**
         * Extracts poems based on significant whitespace breaks
         */
        extractPoemsBySignificantBreaks(htmlParts, filename) { // MANUAL FIX 4
            const poems = [];

            htmlParts.forEach((part, index) => {
                if (!part.trim()) return;

                const tempDiv = document.createElement('div');
                tempDiv.innerHTML = part.trim();
                const content = tempDiv.textContent.trim();

                if (content.length > 20) { // Ensure sufficient content to be a poem
                    const lines = content.split('\n').filter(line => line.trim());
                    const title = lines[0] && lines[0].length < 100 ?
                        lines[0].trim() :
                        `Poem ${index + 1}`;

                    poems.push(this.createPoemObject(title, content, this.preserveFormattingInHtml(part.trim()), filename));
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
            console.log(`Creating single poem object for "${filename}".`);
            const title = this.extractTitle(tempDiv, filename);
            const wordCount = content.split(/\s+/).filter(word => word.length > 0).length;

            return {
                id: Date.now() + Math.random(),
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
            console.log(`Attempting to extract title for "${filename}".`);

            const headings = tempDiv.querySelectorAll('h1, h2, h3');
            for (let i = 0; i < headings.length; i++) {
                const hText = headings[i].textContent.trim();
                if (hText.length > 0 && hText.length < 150) {
                    title = hText;
                    console.log(`  Title found from heading: "${title}"`);
                    break;
                }
            }

            if (!title) {
                const paragraphs = tempDiv.querySelectorAll('p');
                for (let i = 0; i < Math.min(3, paragraphs.length); i++) {
                    const p = paragraphs[i];
                    const pText = p.textContent.trim();
                    if (pText.length > 0 && pText.length < 150) {
                        const isBold = p.querySelector('strong, b') !== null;
                        const isCentered = p.style.textAlign === 'center';

                        if (isBold || isCentered) {
                            title = pText;
                            console.log(`  Title found from bold/centered paragraph: "${title}"`);
                            break;
                        }
                    }
                }
            }

            if (!title) {
                const paragraphs = tempDiv.querySelectorAll('p');
                if (paragraphs.length > 0) {
                    const firstParagraphText = paragraphs[0].textContent.trim();
                    if (firstParagraphText.length > 0) {
                        const firstLine = firstParagraphText.split('\n')[0].trim();
                        if (firstLine.length > 0 && firstLine.length < 150) {
                            title = firstLine;
                            console.log(`  Title found from first line of first paragraph: "${title}"`);
                        }
                    }
                }
            }

            if (!title) {
                title = filename.replace(/\.docx$/i, '').replace(/[_-]/g, ' ').trim();
                console.log(`  Title falling back to cleaned filename: "${title}"`);
            }

            title = title.replace(/\s+/g, ' ').trim();
            if (title.length > 150) {
                title = title.substring(0, 147) + '...';
            }

            if (!title) {
                title = "Untitled Poem";
                console.log(`  Title defaulted to "Untitled Poem".`);
            }
            return title;
        }

        /**
         * Updates the display of loaded poems and their count.
         * Attaches drag-and-drop and button event listeners to each poem element.
         */
        updateDisplay() {
            console.log('Updating display. Current poem count:', this.poems.length);
            const poemList = document.getElementById('poemList');
            const poemCountSpan = document.getElementById('poemCount');
            const downloadBtn = document.getElementById('downloadBtn');
            const clearBtn = document.getElementById('clearBtn');

            if (!poemList || !poemCountSpan || !downloadBtn || !clearBtn) {
                console.error('Required display elements not found for updateDisplay');
                return;
            }

            poemList.innerHTML = '';
            poemCountSpan.textContent = this.poems.length;

            if (this.poems.length === 0) {
                poemList.innerHTML = `
                    <div class="widget-empty-state">
                        <p>No poems loaded yet. Upload and process Word documents to begin.</p>
                    </div>
                `;
                downloadBtn.disabled = true;
                clearBtn.disabled = true;
                console.log('Display updated: No poems loaded, buttons disabled.');
            } else {
                this.poems.forEach((poem, index) => {
                    const poemDiv = this.createPoemElement(poem, index);
                    poemList.appendChild(poemDiv);
                });
                downloadBtn.disabled = false;
                clearBtn.disabled = false;
                console.log('Display updated: Poems rendered, buttons enabled.');
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
            poemDiv.setAttribute('draggable', 'true');
            poemDiv.setAttribute('data-index', index);
            poemDiv.setAttribute('role', 'listitem');
            poemDiv.setAttribute('aria-label', `Poem: ${poem.title}, position ${index + 1} of ${this.poems.length}. Press Ctrl+Up/Down to move, Delete to remove.`);
            poemDiv.setAttribute('tabindex', '0');

            const preview = poem.content.length > 100
                ? poem.content.substring(0, 100).split('\n')[0] + '...'
                : poem.content.split('\n')[0];

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
            poemDiv.addEventListener('dragstart', (e) => {
                this.draggedIndex = index;
                poemDiv.classList.add('dragging');
                e.dataTransfer.effectAllowed = 'move';
                e.dataTransfer.setData('text/plain', index.toString());
                this.announceToScreenReader('process-status', `Started dragging ${this.poems[index].title}`);
                console.log(`Drag started for poem "${this.poems[index].title}" at index ${index}.`);
            });

            poemDiv.addEventListener('dragend', () => {
                document.querySelectorAll('.widget-poem-item').forEach(item => {
                    item.classList.remove('dragging');
                    item.classList.remove('drag-over');
                });
                this.draggedIndex = null;
                console.log('Drag ended.');
            });

            poemDiv.addEventListener('dragover', (e) => {
                e.preventDefault();
                e.dataTransfer.dropEffect = 'move';
                const targetElement = e.currentTarget;
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
                console.log(`Dropped poem from index ${draggedIdx} to index ${dropTargetIndex}.`);
                if (draggedIdx !== dropTargetIndex && !isNaN(draggedIdx)) {
                    this.movePoem(draggedIdx, dropTargetIndex);
                }
            });

            poemDiv.addEventListener('keydown', (e) => {
                if (e.key === 'ArrowUp' && e.ctrlKey && index > 0) {
                    e.preventDefault();
                    console.log(`Keyboard move up for index ${index}.`);
                    this.movePoem(index, index - 1);
                    requestAnimationFrame(() => {
                        const newPoemDiv = document.querySelector(`.widget-poem-item[data-index="${index - 1}"]`);
                        if (newPoemDiv) newPoemDiv.focus();
                        this.announceToScreenReader('process-status', `Moved ${this.poems[index - 1].title} to position ${index}.`);
                    });
                } else if (e.key === 'ArrowDown' && e.ctrlKey && index < this.poems.length - 1) {
                    e.preventDefault();
                    console.log(`Keyboard move down for index ${index}.`);
                    this.movePoem(index, index + 1);
                    requestAnimationFrame(() => {
                        const newPoemDiv = document.querySelector(`.widget-poem-item[data-index="${index + 1}"]`);
                        if (newPoemDiv) newPoemDiv.focus();
                        this.announceToScreenReader('process-status', `Moved ${this.poems[index + 1].title} to position ${index + 2}.`);
                    });
                } else if (e.key === 'Delete' || e.key === 'Backspace') {
                    e.preventDefault();
                    console.log(`Keyboard delete for index ${index}.`);
                    const confirmed = true;
                    if (confirmed) {
                        this.removePoem(index);
                        this.announceToScreenReader('process-status', `Removed ${this.poems[index].title}.`);
                    }
                }
            });

            const moveUpBtn = poemDiv.querySelector('.widget-move-btn:not(.move-down)');
            if (moveUpBtn) {
                moveUpBtn.addEventListener('click', (e) => {
                    e.preventDefault();
                    console.log(`Move up button clicked for index ${index}.`);
                    if (index > 0) {
                        this.movePoem(index, index - 1);
                    }
                });
            }

            const moveDownBtn = poemDiv.querySelector('.widget-move-btn.move-down');
            if (moveDownBtn) {
                moveDownBtn.addEventListener('click', (e) => {
                    e.preventDefault();
                    console.log(`Move down button clicked for index ${index}.`);
                    if (index < this.poems.length - 1) {
                        this.movePoem(index, index + 1);
                    }
                });
            }

            const removeBtn = poemDiv.querySelector('.widget-remove-btn');
            if (removeBtn) {
                removeBtn.addEventListener('click', (e) => {
                    e.preventDefault();
                    console.log(`Remove button clicked for index ${index}.`);
                    const confirmed = true;
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
            console.log(`Moving poem from ${fromIndex} to ${toIndex}.`);
            if (fromIndex < 0 || fromIndex >= this.poems.length ||
                toIndex < 0 || toIndex >= this.poems.length) {
                console.error('Invalid indices for movePoem', fromIndex, toIndex);
                return;
            }

            const [movedPoem] = this.poems.splice(fromIndex, 1);
            this.poems.splice(toIndex, 0, movedPoem);
            this.updateDisplay();
            this.showNotification(`Moved "${movedPoem.title}" from position ${fromIndex + 1} to ${toIndex + 1}`, 'info');
            this.announceToScreenReader('process-status', `Poem moved. New order updated.`);
            console.log(`Poem "${movedPoem.title}" successfully moved.`);
        }

        /**
         * Removes a poem from the array and updates the display.
         * @param {number} index - The index of the poem to remove.
         */
        removePoem(index) {
            console.log(`Removing poem at index ${index}.`);
            if (index < 0 || index >= this.poems.length) {
                console.error('Invalid index for removePoem', index);
                return;
            }
            const removedPoem = this.poems.splice(index, 1)[0];
            this.updateDisplay();
            this.showNotification(`Removed "${removedPoem.title}"`, 'info');
            this.announceToScreenReader('process-status', `Poem ${removedPoem.title} removed.`);
            console.log(`Poem "${removedPoem.title}" successfully removed.`);
        }

        /**
         * Displays a notification message to the user.
         * @param {string} message - The message to display.
         * @param {string} type - The type of notification (success, warning, error, info).
         * @param {number} [duration=5000] - Duration in milliseconds before the notification fades.
         */
        showNotification(message, type, duration = 5000) {
            console.log(`Notification (${type}): ${message}`);
            const notificationContainer = document.getElementById('notificationContainer');
            if (!notificationContainer) {
                console.warn('Notification container not found.');
                return;
            }

            if (this.notificationTimeout) {
                clearTimeout(this.notificationTimeout);
            }
            notificationContainer.innerHTML = '';

            const notificationDiv = document.createElement('div');
            notificationDiv.classList.add('widget-notification', type, 'opacity-0', 'transition-opacity', 'duration-300');
            notificationDiv.setAttribute('role', type === 'error' ? 'alert' : 'status');
            notificationDiv.innerHTML = `
                <span class="mr-2">${this._getNotificationIcon(type)}</span>
                <span>${this.escapeHtml(message)}</span>
            `;
            notificationContainer.appendChild(notificationDiv);

            setTimeout(() => {
                notificationDiv.classList.remove('opacity-0');
            }, 10);

            if (duration > 0) {
                this.notificationTimeout = setTimeout(() => {
                    notificationDiv.classList.add('opacity-0');
                    notificationDiv.addEventListener('transitionend', () => {
                        if (notificationDiv.parentNode) {
                            notificationDiv.parentNode.removeChild(notificationDiv);
                        }
                    }, { once: true });
                }, duration);
            }
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
                console.log(`ARIA announcement for "${elementId}": ${message}`);
            } else {
                console.warn(`ARIA live region element with ID "${elementId}" not found.`);
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
                const poemAnchorId = `poem-${index + 1}-${poem.id}`;
                tocHtml += `<li><a href="#${poemAnchorId}" style="color: #007bff; text-decoration: none;">${this.escapeHtml(poem.title)}</a></li>\n`;
            });
            tocHtml += `</ol>\n\n`;
            tocHtml += `<div style="page-break-after: always;"></div>\n`;
            return tocHtml;
        }

        /**
         * Downloads the combined document in the selected format.
         */
        async downloadCombinedDocument() {
            console.log('Initiating download of combined document.');
            if (this.poems.length === 0) {
                this.showNotification('No poems to download!', 'warning');
                console.warn('Download attempted with no poems.');
                return;
            }

            const exportFormat = document.getElementById('exportFormat').value;
            const downloadBtn = document.getElementById('downloadBtn');

            downloadBtn.disabled = true;
            const originalText = downloadBtn.textContent;
            downloadBtn.textContent = 'Generating...';
            this.showNotification(`Generating ${exportFormat.toUpperCase()}...`, 'info', 0);

            try {
                let filename = 'Combined_Poems';
                let blob;

                const combinedHtml = this._generateHtmlContentForExport();
                console.log(`Generated combined HTML for export. Format: ${exportFormat}`);

                switch (exportFormat) {
                    case 'html':
                        blob = new Blob([combinedHtml], { type: 'text/html' });
                        filename += '.html';
                        break;
                    case 'docx':
                        // Generate MHTML for better Word compatibility
                        const mhtmlContent = this._generateMhtmlContentForExport(combinedHtml);
                        blob = new Blob([mhtmlContent], { type: 'application/x-mimearchive' });
                        filename += '.mht'; // Use .mht extension for MHTML
                        this.showNotification('Downloading as .mht (Web Archive). This format offers better compatibility with Word for HTML content. You may need to "Save As" .docx in Word for full features.', 'info', 10000);
                        break;
                    case 'pdf':
                        console.log('Generating PDF...');
                        console.log('HTML content for PDF:', combinedHtml.substring(0, 500) + '...');
                        await this._generatePdfOutput(combinedHtml, filename);
                        this.showNotification('PDF generated!', 'success');
                        this.resetDownloadUI(downloadBtn, originalText);
                        console.log('PDF generation and download complete.');
                        return;
                    default:
                        this.showNotification('Invalid export format selected.', 'error');
                        console.error('Invalid export format:', exportFormat);
                        return;
                }

                saveAs(blob, filename);
                this.showNotification('Document downloaded successfully!', 'success');
                console.log(`File "${filename}" downloaded.`);

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
            console.log('Resetting download UI.');
            downloadBtn.textContent = originalText;
            downloadBtn.disabled = this.poems.length === 0;
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
                        font-family: 'Times New Roman', serif; /* MANUAL FIX 5 */
                        line-height: 1.6;
                        margin: 40px; /* MANUAL FIX 5 */
                        white-space: pre-wrap; /* MANUAL FIX 5 */
                    }
                    .poem { /* MANUAL FIX 5 */
                        margin-bottom: 50px;
                        page-break-after: auto;
                        white-space: pre-wrap;
                        font-family: 'Courier New', monospace; /* Better for preserving spacing */
                    }
                    .poem-title { /* MANUAL FIX 5 */
                        font-size: 18px;
                        font-weight: bold;
                        text-align: center;
                        margin-bottom: 20px;
                        font-family: 'Times New Roman', serif;
                    }
                    .poem-content { /* MANUAL FIX 5 */
                        white-space: pre-wrap;
                        font-family: 'Courier New', monospace;
                        line-height: 1.4;
                    }
                    .page-break { page-break-before: always; } /* MANUAL FIX 5 */
                    p { margin-bottom: 0; } /* MANUAL FIX 5 */
                    br { line-height: 1.2; } /* MANUAL FIX 5 */
                </style>
            </head>
            <body>
                <h1>A Collection of Poems</h1>
                <div class="table-of-contents">
                    ${this.generateTableOfContentsHtml()}
                </div>
            `;

            this.poems.forEach((poem, index) => {
                const poemAnchorId = `poem-${index + 1}-${poem.id}`;
                combinedHtml += `
                <div class="poem-container" id="${poemAnchorId}">
                    <h2>${this.escapeHtml(poem.title)}</h2>
                    <p class="poem-source">From: ${this.escapeHtml(poem.filename)}</p>
                    ${poem.htmlContent}
                </div>
                `;
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
         * Generates MHTML content from an HTML string for better Word compatibility.
         * @param {string} htmlContent - The HTML string to convert to MHTML.
         * @returns {string} The MHTML string.
         */
        _generateMhtmlContentForExport(htmlContent) {
            const boundary = `----=_NextPart_${Math.random().toString().slice(2)}`;
            let mhtml = `MIME-Version: 1.0\n`;
            mhtml += `Content-Type: multipart/related; boundary="${boundary}"\n\n`;

            mhtml += `--${boundary}\n`;
            mhtml += `Content-Type: text/html; charset="utf-8"\n`;
            mhtml += `Content-Transfer-Encoding: quoted-printable\n`;
            mhtml += `Content-Location: about:blank\n\n`;

            mhtml += this._quotedPrintableEncode(htmlContent) + '\n\n';

            mhtml += `--${boundary}--`;
            return mhtml;
        }

        /**
         * Simple quoted-printable encoder (basic implementation, might need more robust for complex HTML)
         * @param {string} str - The string to encode.
         * @returns {string} The quoted-printable encoded string.
         */
        _quotedPrintableEncode(str) {
            return str.replace(/=/g, '=3D')
                      .replace(/\?/g, '=3F')
                      .replace(/_/g, '=5F')
                      .replace(/\r?\n/g, '=\r\n')
                      .replace(/[\x00-\x1F\x7F-\xFF]/g, (char) => {
                          const byte = char.charCodeAt(0);
                          return '=' + byte.toString(16).toUpperCase().padStart(2, '0');
                      });
        }

        /**
         * Generates and downloads a PDF document from the given HTML content.
         * @param {string} htmlContent - The HTML string to convert to PDF.
         * @param {string} filename - The desired filename for the PDF.
         */
        async _generatePdfOutput(htmlContent, filename) {
            const opt = {
                margin: [20, 20, 20, 20],
                filename: filename,
                image: { type: 'jpeg', quality: 0.98 },
                html2canvas: { scale: 2 },
                jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
            };

            const tempDiv = document.createElement('div');
            tempDiv.innerHTML = htmlContent;
            tempDiv.style.width = '210mm';
            tempDiv.style.margin = '0 auto';
            tempDiv.style.visibility = 'hidden';
            document.body.appendChild(tempDiv);
            console.log('Temporary div created and appended for PDF generation.');

            // Increased delay to 500ms
            await new Promise(resolve => setTimeout(resolve, 500));

            try {
                await html2pdf().set(opt).from(tempDiv).save();
                console.log('html2pdf finished saving.');
            } finally {
                if (tempDiv.parentNode) {
                    document.body.removeChild(tempDiv);
                    console.log('Temporary div removed.');
                }
            }
        }
    }

    document.addEventListener('DOMContentLoaded', () => {
        new PoemCompiler();
    });
})();
