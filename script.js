(function() {
    'use strict';

    class PacketCompiler {
        constructor() {
            this.packets = [];
            this.selectedFiles = [];
            this.draggedIndex = null;
            this.isProcessing = false;
            this.notificationTimeout = null;
            console.log('PacketCompiler initialized.');
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
                this.clearAllPackets();
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
         * Processes the selected Word documents to extract packets.
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
                let processedPacketCount = 0;
                let skippedCount = 0;
                const totalFiles = this.selectedFiles.length;
                console.log(`Processing ${totalFiles} selected files.`);
                const errors = [];

                for (let i = 0; i < this.selectedFiles.length; i++) {
                    const file = this.selectedFiles[i];
                    console.log(`Processing file ${i + 1}/${totalFiles}: ${file.name}`);

                    try {
                        const packetsFromFile = await this.extractPacketsFromDocument(file);
                        console.log(`Extracted ${packetsFromFile ? packetsFromFile.length : 0} potential packets from ${file.name}`);
                        if (packetsFromFile && packetsFromFile.length > 0) {
                            for (const packetData of packetsFromFile) {
                                if (packetData && packetData.content && packetData.content.trim().length > 0) {
                                    // Use a combination of title and content for duplication check
                                    const isDuplicate = this.packets.some(existing =>
                                        existing.title.toLowerCase() === packetData.title.toLowerCase() &&
                                        existing.content.trim() === packetData.content.trim()
                                    );

                                    if (!isDuplicate) {
                                        this.packets.push(packetData);
                                        processedPacketCount++;
                                        console.log(`Added new packet: "${packetData.title}" from "${file.name}"`);
                                    } else {
                                        skippedCount++;
                                        console.warn(`Duplicate packet detected and skipped: "${packetData.title || 'Untitled'}" from "${file.name}"`);
                                    }
                                } else {
                                    console.warn(`Packet data from ${file.name} was empty or invalid.`);
                                }
                            }
                        } else {
                            errors.push(`${file.name}: No valid packets found`);
                            console.warn(`No valid packets found in ${file.name}.`);
                        }
                    } catch (error) {
                        console.error(`Error processing ${file.name}:`, error);
                        errors.push(`${file.name}: ${error.message}`);
                    }

                    const progress = ((i + 1) / totalFiles) * 100;
                    progressBar.style.width = `${progress}%`;
                    progressBar.setAttribute('aria-valuenow', Math.round(progress).toString());
                    console.log(`Progress: ${Math.round(progress)}%`);

                    // Use requestAnimationFrame to ensure UI updates are rendered
                    await new Promise(resolve => requestAnimationFrame(resolve));
                }

                console.log('Finished processing all files. Resetting UI.');
                this.resetProcessingUI();

                if (processedPacketCount > 0) {
                    this.updateDisplay();
                    let message = `Successfully processed ${processedPacketCount} new packet${processedPacketCount > 1 ? 's' : ''}!`;
                    if (skippedCount > 0) {
                        message += ` (${skippedCount} duplicate${skippedCount > 1 ? 's' : ''} skipped)`;
                    }
                    this.showNotification(message, 'success');
                    this.announceToScreenReader('process-status', `${processedPacketCount} packets processed successfully`);
                    this.resetFileInput();
                    console.log('Packet processing complete. Display updated.');
                } else {
                    let message = 'No new packets found in the uploaded documents!';
                    if (skippedCount > 0) {
                        message = `All uploaded packets were duplicates or had no new content.`;
                    }
                    this.showNotification(message, 'warning');
                    this.announceToScreenReader('process-status', 'No new packets found');
                    console.log('No new packets added after processing.');
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
         * Clears all loaded packets and updates the display.
         */
        clearAllPackets() {
            console.log('Clearing all packets.');
            this.packets = [];
            this.updateDisplay();
            this.resetFileInput();
            this.showNotification('All packets cleared!', 'info');
            this.announceToScreenReader('process-status', 'All packets cleared.');
        }

        /**
         * Extracts HTML content from a DOCX file using Mammoth.js
         * and attempts to identify multiple packets within it.
         * @param {File} file - The DOCX file to process.
         * @returns {Promise<Array<Object>>} A promise resolving to an array of packet objects.
         * @throws {Error} If Mammoth.js is not loaded or content extraction fails.
         */
        async extractPacketsFromDocument(file) {
            console.log(`Attempting to extract packets from "${file.name}"...`);
            if (!window.mammoth) {
                console.error('Mammoth library (window.mammoth) is not loaded.');
                throw new Error('Mammoth library not loaded. Please check the script tag.');
            }

            try {
                const arrayBuffer = await file.arrayBuffer();
                console.log(`File "${file.name}" converted to ArrayBuffer.`);
                const result = await window.mammoth.convertToHtml({ arrayBuffer });

                // --- START OF DEBUGGING LINE FOR IMAGE ISSUE ---
                console.log(`Mammoth.js RAW HTML output for "${file.name}":`, result.value);
                // --- END OF DEBUGGING LINE ---

                if (!result.value) {
                    console.warn(`Mammoth returned no HTML content for "${file.name}".`);
                    throw new Error('No content extracted from document by Mammoth.');
                }

                const html = result.value;
                const tempDiv = document.createElement('div');
                tempDiv.innerHTML = html;
                const fullContent = tempDiv.textContent.trim();
                console.log(`Full plain text content length for "${file.name}": ${fullContent.length}`);

                if (!fullContent || fullContent.length < 50) { // Increased minimum length for the whole document
                    console.warn(`Document "${file.name}" appears empty or too short after extraction.`);
                    throw new Error('Document appears to be empty or too short after extraction.');
                }

                // Attempt to identify multiple packets
                const identifiedPackets = this.identifyMultiplePackets(tempDiv, file.name, html);
                console.log(`identifyMultiplePackets returned ${identifiedPackets.length} packets for "${file.name}".`);

                // Crucial Fallback: If multiple packet detection yields 0 or 1 *meaningful* packet,
                // treat the whole document as one. This handles unusually formatted single packets.
                // A "meaningful" packet for this check means substantial content.
                if (identifiedPackets.length <= 1 || (identifiedPackets.length > 1 && identifiedPackets.every(p => p.content.length < 100))) {
                    const singlePacket = this.createSinglePacketFromDocument(tempDiv, file.name, html, fullContent);
                    console.log(`Multi-packet detection found ${identifiedPackets.length} potential segments. Treating "${file.name}" as a single packet: "${singlePacket.title}".`);
                    return [singlePacket];
                }

                return identifiedPackets;

            } catch (error) {
                console.error(`Failed to extract content from "${file.name}":`, error);
                throw new Error(`Failed to extract content from "${file.name}": ${error.message}`);
            }
        }

        /**
         * Attempts to identify and separate multiple packets within an HTML document structure.
         * Uses different strategies (headings, paragraph breaks, explicit separators).
         * @param {HTMLElement} tempDiv - A temporary div containing the document's HTML.
         * @param {string} filename - The original filename.
         * @param {string} fullHtml - The full HTML content from Mammoth.js.
         * @returns {Array<Object>} An array of identified packet objects.
         */
        identifyMultiplePackets(tempDiv, filename, fullHtml) {
            console.log(`Starting identifyMultiplePackets for "${filename}".`);
            let packets = [];
            const MIN_CONTENT_LENGTH_FOR_SPLIT = 100; // A segment must have at least this many chars to be a distinct packet

            // Strategy 1: Split by headings (H1, H2, H3)
            const headings = tempDiv.querySelectorAll('h1, h2, h3');
            if (headings.length > 1) {
                const packetsByHeadings = this.extractPacketsByHeadings(tempDiv, filename, headings);
                // Only consider this a successful multi-packet split if at least two packets are substantial
                if (packetsByHeadings.filter(p => p.content.length >= MIN_CONTENT_LENGTH_FOR_SPLIT).length > 1) {
                    console.log(`Strategy 1 (Headings) found ${packetsByHeadings.length} packets.`);
                    return packetsByHeadings;
                } else {
                    console.log(`Strategy 1 (Headings) found segments, but not enough substantial ones to confirm multiple packets.`);
                }
            }

            // Strategy 2: Split by explicit patterns like "***", "---", or multiple empty paragraphs
            const separatorPatterns = [
                /\n\s*\*{3,}\s*\n/g, // ***
                /\n\s*-{3,}\s*\n/g, // ---
                /\n\s*_{3,}\s*\n/g, // ___
                /\n\s*={3,}\s*\n/g, // ===
                /\n\s*~{3,}\s*\n/g, // ~~~
                /(<p>\s*&nbsp;\s*<\/p>){2,}/g, // Two or more empty paragraphs with &nbsp;
                /(<p>\s*<\/p>){2,}/g // Two or more empty paragraphs
            ];

            for (const pattern of separatorPatterns) {
                // Ensure the pattern actually exists and splits the content into more than one part
                if (fullHtml.match(pattern)) {
                    const partsHtml = fullHtml.split(pattern);
                    // Filter out very short or empty parts that might just be separator artifacts
                    const meaningfulParts = partsHtml.filter(part => {
                        const tempPartDiv = document.createElement('div');
                        tempPartDiv.innerHTML = part;
                        return tempPartDiv.textContent.trim().length >= MIN_CONTENT_LENGTH_FOR_SPLIT;
                    });

                    if (meaningfulParts.length > 1) {
                        const packetsBySeparator = this.extractPacketsBySeparator(meaningfulParts, filename);
                        if (packetsBySeparator.length > 1) {
                            console.log(`Strategy 2 (Separators: ${pattern}) found ${packetsBySeparator.length} packets.`);
                            return packetsBySeparator; // Return early if a clear separator is found
                        }
                    }
                }
            }

            // Strategy 3: Split by significant paragraph breaks (very conservative)
            // This is the most ambiguous strategy, so it should be the last resort and very strict.
            const paragraphs = Array.from(tempDiv.querySelectorAll('p'));
            if (paragraphs.length > 3) { // Need a good number of paragraphs to consider this
                const packetsByParagraphs = this.extractPacketsByParagraphSeparation(tempDiv, filename, paragraphs);
                 if (packetsByParagraphs.filter(p => p.content.length >= MIN_CONTENT_LENGTH_FOR_SPLIT).length > 1) {
                    console.log(`Strategy 3 (Paragraph Separation) found ${packetsByParagraphs.length} packets.`);
                    return packetsByParagraphs;
                } else {
                    console.log(`Strategy 3 (Paragraph Separation) found segments, but not enough substantial ones to confirm multiple packets.`);
                }
            }

            console.log(`No strong multi-packet separation detected for "${filename}".`);
            return []; // Return empty, which will trigger the single-packet fallback
        }

        /**
         * Extracts packets by identifying text blocks separated by heading tags (h1, h2, h3).
         * @param {HTMLElement} tempDiv - The temporary div containing the document HTML.
         * @param {string} filename - The name of the original file.
         * @param {NodeList<HTMLElement>} headings - A NodeList of h1, h2, h3 elements.
         * @returns {Array<Object>} An array of packet objects.
         */
        extractPacketsByHeadings(tempDiv, filename, headings) {
            const packets = [];
            const allElements = Array.from(tempDiv.children);
            console.log(`  Extracting by headings for "${filename}". Found ${headings.length} headings.`);
            const MIN_POEM_LENGTH_HEADING = 50; // Minimum characters for a packet section identified by heading

            for (let i = 0; i < headings.length; i++) {
                const currentHeading = headings[i];
                const nextHeading = headings[i + 1];

                const title = currentHeading.textContent.trim();
                // Ensure heading is meaningful
                if (title.length === 0 || title.length > 200 || title.split(/\s+/).length > 20) { // Max 20 words for a title
                    console.log(`    Skipping heading with invalid title: "${title}"`);
                    continue;
                }

                const startIndex = allElements.indexOf(currentHeading);
                const endIndex = nextHeading ? allElements.indexOf(nextHeading) : allElements.length;

                // Collect all elements between current heading and the next (or end of document)
                const packetElements = allElements.slice(startIndex + 1, endIndex);

                // Filter out any empty text nodes or very short paragraphs that might be artifacts
                const meaningfulElements = packetElements.filter(el => {
                    // Check if element has actual content or is a line break (<br>) or non-empty paragraph
                    // Consider content beyond just whitespace or &nbsp;
                    return el.textContent.trim().length > 0 || (el.tagName === 'P' && el.innerHTML.trim() !== '&nbsp;' && el.innerHTML.trim() !== '');
                });

                if (meaningfulElements.length === 0) {
                    console.log(`    Skipping heading "${title}" as no meaningful content found before next heading/end.`);
                    continue;
                }

                // Preserve the structure and formatting as much as possible for htmlContent
                const packetHtml = meaningfulElements.map(el => el.outerHTML).join('\n');
                const packetContent = meaningfulElements.map(el => {
                    // Convert <br> to newline, otherwise use textContent.
                    // This is for plain text content for word count and duplication check.
                    return el.tagName === 'BR' ? '\n' : el.textContent;
                }).join('\n').trim();

                // Add the heading itself to the packet's htmlContent to retain its styling
                const fullPacketHtml = currentHeading.outerHTML + '\n' + packetHtml;

                if (packetContent.length >= MIN_POEM_LENGTH_HEADING) {
                    packets.push(this.createPacketObject(title, packetContent, fullPacketHtml, filename));
                } else {
                    console.log(`    Skipping heading "${title}" due to insufficient content (${packetContent.length} chars).`);
                }
            }
            console.log(`  Finished heading extraction. Found ${packets.length} packets.`);
            return packets;
        }

        /**
         * Extracts packets by identifying blocks of paragraphs separated by empty or very short paragraphs.
         * This strategy is now very conservative.
         * @param {HTMLElement} tempDiv - The temporary div containing the document HTML.
         * @param {string} filename - The name of the original file.
         * @param {NodeList<HTMLElement>} paragraphs - A NodeList of paragraph elements.
         * @returns {Array<Object>} An array of packet objects.
         */
        extractPacketsByParagraphSeparation(tempDiv, filename, paragraphs) {
            const packets = [];
            let currentPacketElements = [];
            let currentTitle = '';
            let packetIndex = 1;
            console.log(`  Extracting by paragraph separation for "${filename}". Found ${paragraphs.length} paragraphs.`);
            const MIN_POEM_LENGTH_PARA_SPLIT = 150; // Higher minimum for this ambiguous splitting method

            for (let i = 0; i < paragraphs.length; i++) {
                const p = paragraphs[i];
                const text = p.textContent.trim();
                const html = p.outerHTML;

                // A much stricter heuristic for a "significant break" that indicates a new packet.
                // Require a truly empty paragraph or a paragraph with only non-breaking spaces,
                // OR at least two consecutive empty-like paragraphs.
                const isTrulyEmpty = text.length === 0 || p.innerHTML.trim() === '&nbsp;' || p.innerHTML.trim() === '<br>' || p.innerHTML.trim() === '<br />';
                const isSignificantBreak = isTrulyEmpty && (
                    (i + 1 < paragraphs.length && (paragraphs[i + 1].textContent.trim().length === 0 || paragraphs[i + 1].innerHTML.trim() === '&nbsp;')) ||
                    (i + 2 < paragraphs.length && (paragraphs[i + 2].textContent.trim().length === 0 || paragraphs[i + 2].innerHTML.trim() === '&nbsp;'))
                );

                // Heuristic for what might be a title: short, possibly bold/centered, starts with a capital letter
                // Be very conservative: must be short AND either bold, centered, or ALL CAPS
                const mightBeTitle = text.length > 0 && text.length < 100 &&
                    (p.querySelector('strong, b') || p.style.textAlign === 'center' || (text === text.toUpperCase() && text.length < 50 && text.split(/\s+/).length < 10)); // Max 10 words for an ALL CAPS title

                if (isSignificantBreak) {
                    if (currentPacketElements.length > 0) {
                        const packetContent = currentPacketElements.map(el => el.textContent).join('\n').trim();
                        const packetHtml = currentPacketElements.map(el => el.outerHTML).join('\n');
                        const title = currentTitle || `Packet ${packetIndex} from ${filename}`;

                        if (packetContent.length >= MIN_POEM_LENGTH_PARA_SPLIT) { // Ensure substantial content
                            packets.push(this.createPacketObject(title, packetContent, packetHtml, filename));
                            packetIndex++;
                            console.log(`    Packet #${packetIndex - 1} identified by significant paragraph break: "${title}"`);
                        } else {
                            console.log(`    Skipping short packet segment before break (length: ${packetContent.length}). Likely part of previous/next.`);
                        }
                        currentPacketElements = [];
                        currentTitle = '';
                    }
                    // Skip the actual empty/break paragraphs from being added to any packet's content
                } else if (mightBeTitle && currentPacketElements.length === 0) {
                    // If it's a potential title and no current packet content, start a new one
                    currentTitle = text;
                    currentPacketElements.push(p);
                    console.log(`    Potential title detected: "${text}"`);
                } else {
                    currentPacketElements.push(p);
                }
            }

            // Add any remaining packet elements after the loop finishes
            if (currentPacketElements.length > 0) {
                const packetContent = currentPacketElements.map(el => el.textContent).join('\n').trim();
                const packetHtml = currentPacketElements.map(el => el.outerHTML).join('\n');
                const title = currentTitle || `Packet ${packetIndex} from ${filename}`;

                if (packetContent.length >= MIN_POEM_LENGTH_PARA_SPLIT) {
                    packets.push(this.createPacketObject(title, packetContent, packetHtml, filename));
                    console.log(`    Last packet identified: "${title}"`);
                } else {
                    console.log(`    Skipping last packet segment due to insufficient content (length: ${packetContent.length}). Likely part of previous/next.`);
                }
            }
            console.log(`  Finished paragraph separation. Found ${packets.length} packets.`);
            return packets;
        }

        /**
         * Extracts packets by identifying blocks of HTML content based on detected separator patterns.
         * @param {Array<string>} htmlParts - Array of HTML strings separated by a pattern.
         * @param {string} filename - The name of the original file.
         * @returns {Array<Object>} An array of packet objects.
         */
        extractPacketsBySeparator(htmlParts, filename) {
            const packets = [];
            console.log(`  Extracting by custom separators for "${filename}". Found ${htmlParts.length} parts.`);
            const MIN_POEM_LENGTH_SEPARATOR = 50; // Minimum characters for a packet section identified by separator

            htmlParts.forEach((part, index) => {
                const tempDiv = document.createElement('div');
                tempDiv.innerHTML = part.trim();
                const content = tempDiv.textContent.trim();

                // If the part is just the separator itself or very short, skip it
                if (content.length < MIN_POEM_LENGTH_SEPARATOR && !tempDiv.querySelector('p, h1, h2, h3, h4, h5, h6, pre')) {
                    console.log(`    Skipping part ${index + 1} due to insufficient content after separator.`);
                    return;
                }

                // Attempt to find a title within this part, prioritizing headings or bold/centered text
                let title = '';
                const headings = tempDiv.querySelectorAll('h1, h2, h3');
                if (headings.length > 0) {
                    title = headings[0].textContent.trim();
                } else {
                    const paragraphs = tempDiv.querySelectorAll('p');
                    if (paragraphs.length > 0) {
                        const firstParaText = paragraphs[0].textContent.trim();
                        if (firstParaText.length > 0 && firstParaText.length < 150) {
                            const isBold = paragraphs[0].querySelector('strong, b') !== null;
                            const isCentered = paragraphs[0].style.textAlign === 'center';
                            if (isBold || isCentered) {
                                title = firstParaText;
                            }
                        }
                    }
                }

                if (!title) {
                    // Fallback to first non-empty line as title if no clear title found
                    const lines = content.split('\n').map(line => line.trim()).filter(line => line.length > 0);
                    const firstLine = lines[0] || '';
                    if (firstLine.length > 0 && firstLine.length < 100) {
                        title = firstLine;
                    } else {
                        title = `Packet ${index + 1} from ${filename}`;
                    }
                }

                if (content.length >= MIN_POEM_LENGTH_SEPARATOR) {
                    packets.push(this.createPacketObject(title, content, part.trim(), filename));
                    console.log(`    Packet #${index + 1} identified by separator: "${title}"`);
                } else {
                    console.log(`    Skipping segment after separator due to insufficient content (length: ${content.length}).`);
                }
            });
            console.log(`  Finished separator extraction. Found ${packets.length} packets.`);
            return packets;
        }


        /**
         * Creates a single packet object from an entire document when multiple packets are not detected.
         * @param {HTMLElement} tempDiv - The temporary div containing the document HTML.
         * @param {string} filename - The original filename.
         * @param {string} html - The full HTML content from Mammoth.js.
         * @param {string} content - The full plain text content of the document.
         * @returns {Object} A single packet object.
         */
        createSinglePacketFromDocument(tempDiv, filename, html, content) {
            console.log(`Creating single packet object for "${filename}".`);
            const title = this.extractTitle(tempDiv, filename);
            const wordCount = content.split(/\s+/).filter(word => word.length > 0).length;

            return {
                id: Date.now() + Math.random(),
                title: title,
                content: content,
                htmlContent: html, // The full HTML of the document as a single packet
                filename: filename,
                wordCount: wordCount,
                dateAdded: new Date().toISOString()
            };
        }

        /**
         * Creates a packet object with all necessary properties.
         * @param {string} title - The title of the packet.
         * @param {string} content - The plain text content of the packet.
         * @param {string} htmlContent - The HTML content of the packet.
         * @param {string} filename - The original filename from which the packet was extracted.
         * @returns {Object} The packet object.
         */
        createPacketObject(title, content, htmlContent, filename) {
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

            // Prioritize actual heading tags
            const headings = tempDiv.querySelectorAll('h1, h2, h3');
            for (let i = 0; i < headings.length; i++) {
                const hText = headings[i].textContent.trim();
                // Ensure heading is meaningful (not too short/long, not just numbers etc)
                if (hText.length > 0 && hText.length < 150 && hText.split(/\s+/).length < 25 && /[a-zA-Z]/.test(hText)) {
                    title = hText;
                    console.log(`  Title found from heading: "${title}"`);
                    break;
                }
            }

            // If no heading, check first few paragraphs for bold/centered text
            if (!title) {
                const paragraphs = tempDiv.querySelectorAll('p');
                for (let i = 0; i < Math.min(3, paragraphs.length); i++) {
                    const p = paragraphs[i];
                    const pText = p.textContent.trim();
                    if (pText.length > 0 && pText.length < 150) {
                        const isBold = p.querySelector('strong, b') !== null;
                        const isCentered = p.style.textAlign === 'center';
                        const isAllCaps = pText === pText.toUpperCase() && pText.length < 50 && pText.split(/\s+/).length < 10;

                        if (isBold || isCentered || isAllCaps) {
                            title = pText;
                            console.log(`  Title found from bold/centered/all caps paragraph: "${title}"`);
                            break;
                        }
                    }
                }
            }

            // As a last resort, use the first non-empty line of content or filename
            if (!title) {
                const paragraphs = tempDiv.querySelectorAll('p');
                if (paragraphs.length > 0) {
                    const firstMeaningfulParagraph = Array.from(paragraphs).find(p => p.textContent.trim().length > 0);
                    if (firstMeaningfulParagraph) {
                        const firstParagraphText = firstMeaningfulParagraph.textContent.trim();
                        if (firstParagraphText.length > 0) {
                            const firstLine = firstParagraphText.split('\n')[0].trim();
                            if (firstLine.length > 0 && firstLine.length < 150 && firstLine.split(/\s+/).length < 25) {
                                title = firstLine;
                                console.log(`  Title found from first line of first meaningful paragraph: "${title}"`);
                            }
                        }
                    }
                }
            }

            if (!title) {
                title = filename.replace(/\.docx$/i, '').replace(/[_-]/g, ' ').trim();
                console.log(`  Title falling back to cleaned filename: "${title}"`);
            }

            // Capitalize first letter of fallback titles for better display
            if (title.startsWith('packet ') && title.toLowerCase().includes('from')) {
                 // Leave "Packet X from file" as is
            } else if (title.length > 0) {
                title = title.charAt(0).toUpperCase() + title.slice(1);
            }

            return title;
        }

        /**
         * Updates the display of loaded packets in the UI.
         * Enables/disables the download button based on packet count.
         * Re-initializes drag-and-drop for packet reordering.
         */
        updateDisplay() {
            console.log('Updating display for packets. Total packets:', this.packets.length);
            const packetsList = document.getElementById('packetsList');
            const downloadBtn = document.getElementById('downloadBtn');
            const clearBtn = document.getElementById('clearBtn');
            const placeholder = document.getElementById('packetsPlaceholder');
            const packetCountSpan = document.getElementById('packetCount');

            if (!packetsList || !downloadBtn || !clearBtn || !placeholder || !packetCountSpan) {
                console.error('Required DOM elements for display update not found. Ensure all IDs are correct in HTML.');
                return;
            }

            packetsList.innerHTML = ''; // Clear existing list
            packetCountSpan.textContent = this.packets.length.toString();

            if (this.packets.length === 0) {
                placeholder.style.display = 'block';
                packetsList.style.display = 'none';
                downloadBtn.disabled = true;
                clearBtn.disabled = true;
                this.announceToScreenReader('packet-list-status', 'No packets loaded.');
                console.log('No packets to display. Placeholder shown, buttons disabled.');
                return;
            }

            placeholder.style.display = 'none';
            packetsList.style.display = 'block';
            downloadBtn.disabled = false;
            clearBtn.disabled = false;

            // Re-render packets based on the current order in this.packets array
            this.packets.forEach((packet, index) => {
                const li = document.createElement('li');
                li.className = 'packet-item bg-white p-4 shadow-sm rounded-lg flex items-center justify-between transition-all duration-200 ease-in-out';
                li.draggable = true;
                li.dataset.id = packet.id;
                li.dataset.index = index; // Important for reordering

                li.innerHTML = `
                    <div class="flex-1 min-w-0">
                        <h3 class="text-lg font-semibold text-gray-800 truncate">${this.escapeHtml(packet.title)}</h3>
                        <p class="text-sm text-gray-500 truncate">${this.escapeHtml(packet.filename)} - ${packet.wordCount} words</p>
                    </div>
                    <div class="flex items-center space-x-2 ml-4">
                        <button type="button" class="move-up-btn p-2 rounded-full text-gray-600 hover:bg-gray-200 focus:outline-none focus:ring-2 focus:ring-gray-500 focus:ring-opacity-50" aria-label="Move packet ${this.escapeHtml(packet.title)} up" data-id="${packet.id}" ${index === 0 ? 'disabled' : ''}>
                            <svg xmlns="http://www.w3.org/2000/svg" class="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2">
                                <path stroke-linecap="round" stroke-linejoin="round" d="M5 10l7-7m0 0l7 7m-7-7v18" />
                            </svg>
                        </button>
                        <button type="button" class="move-down-btn p-2 rounded-full text-gray-600 hover:bg-gray-200 focus:outline-none focus:ring-2 focus:ring-gray-500 focus:ring-opacity-50" aria-label="Move packet ${this.escapeHtml(packet.title)} down" data-id="${packet.id}" ${index === this.packets.length - 1 ? 'disabled' : ''}>
                            <svg xmlns="http://www.w3.org/2000/svg" class="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2">
                                <path stroke-linecap="round" stroke-linejoin="round" d="M19 14l-7 7m0 0l-7-7m7 7V3" />
                            </svg>
                        </button>
                        <button type="button" class="view-packet-btn p-2 rounded-full text-blue-600 hover:bg-blue-100 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-opacity-50" aria-label="View packet ${this.escapeHtml(packet.title)}" data-id="${packet.id}">
                            <svg xmlns="http://www.w3.org/2000/svg" class="h-5 w-5" viewBox="0 0 20 20" fill="currentColor">
                                <path d="M10 12a2 2 0 100-4 2 2 0 000 4z" />
                                <path fill-rule="evenodd" d="M.458 10C1.732 5.943 5.522 3 10 3s8.268 2.943 9.542 7c-1.274 4.057-5.064 7-9.542 7S1.732 14.057.458 10zM14 10a4 4 0 11-8 0 4 4 0 018 0z" clip-rule="evenodd" />
                            </svg>
                        </button>
                        <button type="button" class="remove-packet-btn p-2 rounded-full text-red-600 hover:bg-red-100 focus:outline-none focus:ring-2 focus:ring-red-500 focus:ring-opacity-50" aria-label="Remove packet ${this.escapeHtml(packet.title)}" data-id="${packet.id}">
                            <svg xmlns="http://www.w3.org/2000/svg" class="h-5 w-5" viewBox="0 0 20 20" fill="currentColor">
                                <path fill-rule="evenodd" d="M9 2a1 1 0 00-.894.553L7.382 4H4a1 1 0 000 2v10a2 2 0 002 2h8a2 2 0 002-2V6a1 1 0 100-2h-3.382l-.724-1.447A1 1 0 0011 2H9zM7 8a1 1 0 012 0v6a1 1 0 11-2 0V8zm6 0a1 1 0 012 0v6a1 1 0 11-2 0V8z" clip-rule="evenodd" />
                            </svg>
                        </button>
                    </div>
                `;
                packetsList.appendChild(li);
            });

            this.addPacketListEventListeners();
            this.announceToScreenReader('packet-list-status', `${this.packets.length} packets loaded. Use drag and drop or arrows to reorder.`);
            console.log('Packet list rendered and event listeners added.');
        }

        /**
         * Adds event listeners for viewing, removing, and reordering packets.
         */
        addPacketListEventListeners() {
            const packetsList = document.getElementById('packetsList');
            if (!packetsList) {
                console.error('Packet list element not found for event listeners.');
                return;
            }

            // Delegated event listeners for buttons for better performance and dynamic content
            packetsList.removeEventListener('click', this._packetListClickHandler); // Remove old handler if exists
            this._packetListClickHandler = (e) => { // Store handler for removal
                const button = e.target.closest('button');
                if (!button) return;

                const id = button.dataset.id;
                if (button.classList.contains('view-packet-btn')) {
                    this.viewPacket(id);
                } else if (button.classList.contains('remove-packet-btn')) {
                    this.removePacket(id);
                } else if (button.classList.contains('move-up-btn')) {
                    this.movePacketUp(id);
                } else if (button.classList.contains('move-down-btn')) {
                    this.movePacketDown(id);
                }
            };
            packetsList.addEventListener('click', this._packetListClickHandler);


            // Drag and Drop for reordering
            packetsList.removeEventListener('dragstart', this._dragStartHandler);
            packetsList.removeEventListener('dragover', this._dragOverHandler);
            packetsList.removeEventListener('drop', this._dropHandler);
            packetsList.removeEventListener('dragend', this._dragEndHandler);

            this._dragStartHandler = (e) => {
                const target = e.target.closest('.packet-item');
                if (target) {
                    this.draggedIndex = parseInt(target.dataset.index, 10);
                    e.dataTransfer.effectAllowed = 'move';
                    e.dataTransfer.setData('text/plain', this.draggedIndex); // Set data for Firefox compatibility
                    setTimeout(() => target.classList.add('dragging'), 0); // Add class after a tiny delay
                    console.log('Drag started for index:', this.draggedIndex);
                }
            };
            packetsList.addEventListener('dragstart', this._dragStartHandler);

            this._dragOverHandler = (e) => {
                this.preventDefaults(e); // Allow drop
                const target = e.target.closest('.packet-item');
                if (target && target.dataset.index !== undefined && this.draggedIndex !== null) {
                    const dragOverIndex = parseInt(target.dataset.index, 10);
                    const draggedEl = packetsList.querySelector('.packet-item.dragging');

                    if (draggedEl && this.draggedIndex !== dragOverIndex) {
                        const currentParent = target.parentNode;
                        if (currentParent && draggedEl.parentNode === currentParent) {
                            const targetRect = target.getBoundingClientRect();
                            const mouseY = e.clientY;
                            const targetMidY = targetRect.top + targetRect.height / 2;

                            if (mouseY < targetMidY && draggedEl !== target.previousElementSibling) {
                                currentParent.insertBefore(draggedEl, target);
                            } else if (mouseY >= targetMidY && draggedEl !== target.nextElementSibling) {
                                currentParent.insertBefore(draggedEl, target.nextSibling);
                            }
                        }
                    }
                }
            };
            packetsList.addEventListener('dragover', this._dragOverHandler);

            // No specific action needed for dragleave for this simple reorder logic

            this._dropHandler = (e) => {
                this.preventDefaults(e);
                const draggedEl = packetsList.querySelector('.packet-item.dragging'); // Get the element still marked as dragging

                if (draggedEl && this.draggedIndex !== null) {
                    const newIndex = Array.from(packetsList.children).indexOf(draggedEl); // Get the new visual index

                    if (this.draggedIndex !== newIndex && newIndex !== -1) {
                        console.log(`Drop detected. Original Dragged Index: ${this.draggedIndex}, New Visual Index: ${newIndex}`);

                        const [draggedPacket] = this.packets.splice(this.draggedIndex, 1);
                        this.packets.splice(newIndex, 0, draggedPacket);

                        this.showNotification(`Reordered packet "${draggedPacket.title}"`, 'info');
                        this.announceToScreenReader('packet-list-status', `Packet ${draggedPacket.title} moved to position ${newIndex + 1}.`);

                        this.draggedIndex = null; // Reset
                        draggedEl.classList.remove('dragging'); // Remove dragging class

                        // Re-render the list to reflect the new order and update data-index attributes correctly
                        this.updateDisplay();
                    } else {
                        console.log('Packet dropped on its original position or invalid drop. No reordering needed.');
                        this.draggedIndex = null;
                        if (draggedEl) draggedEl.classList.remove('dragging');
                    }
                } else if (draggedEl) {
                    // If dropped outside a valid target, just remove dragging class
                    draggedEl.classList.remove('dragging');
                    this.draggedIndex = null;
                }
            };
            packetsList.addEventListener('drop', this._dropHandler);

            this._dragEndHandler = (e) => {
                const draggedEl = packetsList.querySelector('.packet-item.dragging');
                if (draggedEl) {
                    draggedEl.classList.remove('dragging');
                }
                this.draggedIndex = null; // Reset
                console.log('Drag ended. draggedIndex reset.');
            };
            packetsList.addEventListener('dragend', this._dragEndHandler);

            console.log('Drag and drop listeners (re)added to packet list.');
        }

        /**
         * Moves a packet up in the list (towards the beginning of the array).
         * @param {string} id - The ID of the packet to move.
         */
        movePacketUp(id) {
            const index = this.packets.findIndex(p => p.id == id);
            if (index > 0) {
                const [packet] = this.packets.splice(index, 1);
                this.packets.splice(index - 1, 0, packet);
                this.updateDisplay();
                this.showNotification(`Moved "${packet.title}" up.`, 'info');
                this.announceToScreenReader('packet-list-status', `Packet ${packet.title} moved up to position ${index}.`);
                // Re-focus the moved packet's up button for better accessibility
                document.querySelector(`li[data-id="${id}"] .move-up-btn`)?.focus();
            } else {
                this.showNotification('Packet is already at the top.', 'info');
            }
        }

        /**
         * Moves a packet down in the list (towards the end of the array).
         * @param {string} id - The ID of the packet to move.
         */
        movePacketDown(id) {
            const index = this.packets.findIndex(p => p.id == id);
            if (index < this.packets.length - 1 && index !== -1) {
                const [packet] = this.packets.splice(index, 1);
                this.packets.splice(index + 1, 0, packet);
                this.updateDisplay();
                this.showNotification(`Moved "${packet.title}" down.`, 'info');
                this.announceToScreenReader('packet-list-status', `Packet ${packet.title} moved down to position ${index + 2}.`);
                // Re-focus the moved packet's down button for better accessibility
                document.querySelector(`li[data-id="${id}"] .move-down-btn`)?.focus();
            } else {
                this.showNotification('Packet is already at the bottom.', 'info');
            }
        }

        /**
         * Removes a packet from the list by its ID.
         * @param {string} id - The ID of the packet to remove.
         */
        removePacket(id) {
            console.log('Attempting to remove packet with ID:', id);
            const initialCount = this.packets.length;
            const removedPacket = this.packets.find(packet => packet.id == id);
            this.packets = this.packets.filter(packet => packet.id != id); // Use != for loose comparison with dataset.id (string)
            if (this.packets.length < initialCount) {
                this.showNotification(`Packet "${removedPacket ? removedPacket.title : 'Unknown'}" removed!`, 'success');
                this.announceToScreenReader('packet-list-status', `Packet ${removedPacket ? removedPacket.title : 'Unknown'} removed.`);
                this.updateDisplay();
                console.log('Packet removed successfully. Remaining packets:', this.packets.length);
            } else {
                console.warn('Packet with ID not found:', id);
            }
        }

        /**
         * Displays a packet's content in a modal.
         * @param {string} id - The ID of the packet to view.
         */
        viewPacket(id) {
            console.log('Viewing packet with ID:', id);
            const packet = this.packets.find(p => p.id == id);
            if (packet) {
                const modal = document.getElementById('packetModal');
                const modalTitle = document.getElementById('packetModalTitle');
                const modalContent = document.getElementById('packetModalContent');
                const closeModalBtn = document.getElementById('closeModal');

                if (!modal || !modalTitle || !modalContent || !closeModalBtn) {
                    console.error('Modal elements not found.');
                    this.showNotification('Error: Modal display elements missing.', 'error');
                    return;
                }

                modalTitle.textContent = packet.title;
                // Use innerHTML to preserve formatting from Mammoth.js
                modalContent.innerHTML = packet.htmlContent;
                modal.classList.remove('hidden');
                modal.setAttribute('aria-hidden', 'false');
                modal.focus(); // Focus the modal for accessibility

                const closeHandler = () => {
                    modal.classList.add('hidden');
                    modal.setAttribute('aria-hidden', 'true');
                    // Return focus to the button that opened the modal if possible
                    document.querySelector(`button[data-id="${id}"]`)?.focus();
                    closeModalBtn.removeEventListener('click', closeHandler); // Clean up listener
                    document.removeEventListener('keydown', handleEscape);
                };

                closeModalBtn.addEventListener('click', closeHandler);

                // Close modal on escape key
                const handleEscape = (e) => {
                    if (e.key === 'Escape') {
                        closeHandler();
                    }
                };
                document.addEventListener('keydown', handleEscape);

                console.log(`Modal opened for "${packet.title}".`);
            } else {
                this.showNotification('Packet not found.', 'error');
                console.warn('Attempted to view non-existent packet ID:', id);
            }
        }

        /**
         * Combines all loaded packets into a single HTML document and triggers a download.
         */
        downloadCombinedDocument() {
            if (this.packets.length === 0) {
                this.showNotification('No packets to download!', 'warning');
                console.warn('Download attempted with no packets.');
                return;
            }

            console.log('Preparing combined HTML document for download.');

            // Generate Table of Contents
            let tableOfContentsHtml = '';
            if (this.packets.length > 0) {
                tableOfContentsHtml = `
        <nav class="table-of-contents">
            <h2>Table of Contents</h2>
            <ol>
`;
                this.packets.forEach(packet => {
                    tableOfContentsHtml += `
                <li><a href="#packet-${packet.id}">${this.escapeHtml(packet.title)}</a></li>
`;
                });
                tableOfContentsHtml += `
            </ol>
        </nav>
`;
            }


            let combinedHtml = `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Combined Packets</title>
    <style>
        body { font-family: sans-serif; line-height: 1.6; max-width: 800px; margin: 2em auto; padding: 0 1em; color: #333; }
        h1, h2, h3 { color: #2c3e50; margin-top: 1.5em; margin-bottom: 0.5em; }
        h1 { font-size: 2.2em; text-align: center; border-bottom: 2px solid #eee; padding-bottom: 0.5em; }
        h2 { font-size: 1.8em; }
        h3 { font-size: 1.4em; }
        .packet-section { margin-bottom: 2em; padding-bottom: 1em; border-bottom: 1px dashed #eee; }
        .packet-section:last-child { border-bottom: none; margin-bottom: 0; padding-bottom: 0; }
        /* Mammoth.js often wraps content in paragraphs, so default to no top margin */
        .packet-content p { margin-top: 0; margin-bottom: 0.5em; }
        /* Ensure line breaks are visible if they are represented as <br> */
        .packet-content br {
            display: block;
            content: "";
            margin-top: 0.5em; /* Adds vertical space for line breaks */
        }
        .packet-content pre { white-space: pre-wrap; word-wrap: break-word; font-family: monospace; }
        .packet-content code { white-space: pre-wrap; word-wrap: break-word; font-family: monospace; }
        /* Basic alignment from Mammoth.js output */
        p[align="center"] { text-align: center; }
        p[align="right"] { text-align: right; }
        /* Preserve white-space for pre-formatted packet lines */
        .packet-content pre { white-space: pre-wrap; word-wrap: break-word; }
        .packet-content code { white-space: pre-wrap; word-wrap: break-word; }
        /* Ensure other block elements like div maintain spacing */
        .packet-content > div { margin-bottom: 0.5em; }
        /* Basic image styling */
        .packet-content img { max-width: 100%; height: auto; display: block; margin: 0.5em auto; }

        /* Table of Contents Styling */
        .table-of-contents {
            margin: 2em 0;
            padding: 1em;
            border: 1px solid #eee;
            background-color: #f9f9f9;
            border-radius: 5px;
        }
        .table-of-contents h2 {
            margin-top: 0;
            font-size: 1.5em;
            text-align: center;
            border-bottom: 1px solid #ddd;
            padding-bottom: 0.5em;
            margin-bottom: 1em;
        }
        .table-of-contents ol {
            list-style: decimal;
            padding-left: 2em;
        }
        .table-of-contents li {
            margin-bottom: 0.5em;
        }
        .table-of-contents a {
            color: #3b82f6;
            text-decoration: none;
        }
        .table-of-contents a:hover {
            text-decoration: underline;
        }
    </style>
</head>
<body>
    <h1>Combined Packets</h1>
    ${tableOfContentsHtml}
    <div class="packets-container">
`;

            this.packets.forEach((packet, index) => {
                let cleanedHtmlContent = packet.htmlContent;
                const tempContentDiv = document.createElement('div');
                tempContentDiv.innerHTML = packet.htmlContent;

                // Check if the packet's title (or a very similar version) is the first heading in its htmlContent
                const firstHeading = tempContentDiv.querySelector('h1, h2, h3');
                if (firstHeading) {
                    const normalizedTitle = packet.title.toLowerCase().replace(/[^a-z0-9]/g, '');
                    const normalizedHeadingText = firstHeading.textContent.trim().toLowerCase().replace(/[^a-z0-9]/g, '');

                    // Use a more robust check for similarity
                    // Check if the heading text is a significant portion of the packet title or vice-versa
                    const isSimilar = normalizedTitle.includes(normalizedHeadingText) || normalizedHeadingText.includes(normalizedTitle);
                    const isExactMatch = normalizedTitle === normalizedHeadingText;

                    if (isExactMatch || (isSimilar && firstHeading.textContent.trim().length > 0)) {
                        console.log(`Removing duplicate heading "${firstHeading.textContent.trim()}" from packet "${packet.title}" for export.`);
                        firstHeading.remove();
                        cleanedHtmlContent = tempContentDiv.innerHTML;
                    }
                }

                combinedHtml += `
        <div class="packet-section" id="packet-${packet.id}">
            <h2>${this.escapeHtml(packet.title)}</h2>
            <p class="packet-metadata"><em>Source: ${this.escapeHtml(packet.filename)} | Words: ${packet.wordCount}</em></p>
            <div class="packet-content">
                ${cleanedHtmlContent}
            </div>
        </div>
`;
            });

            combinedHtml += `
    </div>
</body>
</html>`;

            const blob = new Blob([combinedHtml], { type: 'text/html;charset=utf-8' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'Combined_Packets.html'; // Ensure it's an HTML file
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
            this.showNotification('Combined document downloaded as HTML!', 'success');
            this.announceToScreenReader('process-status', 'Combined document downloaded as HTML.');
            console.log('Combined document download initiated.');
        }

        /**
         * Displays a temporary notification message to the user.
         * @param {string} message - The message to display.
         * @param {string} type - The type of notification (e.g., 'success', 'error', 'info', 'warning').
         * @param {number} [duration=5000] - How long the notification should be visible in milliseconds.
         */
        showNotification(message, type, duration = 5000) {
            const notification = document.getElementById('notification');
            if (!notification) {
                console.error('Notification element not found.');
                return;
            }

            // Clear any existing timeout to prevent rapid notifications from hiding too quickly
            if (this.notificationTimeout) {
                clearTimeout(this.notificationTimeout);
                this.notificationTimeout = null;
            }

            notification.textContent = message;
            notification.className = `notification fixed bottom-4 right-4 p-3 rounded-md shadow-lg text-white opacity-0 transition-opacity duration-300 z-50`;

            switch (type) {
                case 'success':
                    notification.classList.add('bg-green-500');
                    break;
                case 'error':
                    notification.classList.add('bg-red-500');
                    break;
                case 'info':
                    notification.classList.add('bg-blue-500');
                    break;
                case 'warning':
                    notification.classList.add('bg-yellow-500');
                    notification.classList.add('text-gray-900'); // Ensure text is visible on yellow
                    break;
                default:
                    notification.classList.add('bg-gray-700');
            }

            // Show notification
            requestAnimationFrame(() => {
                notification.classList.remove('opacity-0');
                notification.classList.add('opacity-100');
            });

            // Hide after duration
            this.notificationTimeout = setTimeout(() => {
                notification.classList.remove('opacity-100');
                notification.classList.add('opacity-0');
                this.notificationTimeout = null;
            }, duration);

            console.log(`Notification: ${message} (${type})`);
        }

        /**
         * Announces messages to screen readers using an ARIA live region.
         * @param {string} regionId - The ID of the live region element.
         * @param {string} message - The message to announce.
         */
        announceToScreenReader(regionId, message) {
            const liveRegion = document.getElementById(regionId);
            if (liveRegion) {
                liveRegion.textContent = message;
                console.log(`Announced to screen reader (${regionId}): ${message}`);
            } else {
                console.warn(`ARIA live region with ID "${regionId}" not found.`);
            }
        }

        /**
         * Escapes HTML entities in a string to prevent XSS.
         * @param {string} str - The string to escape.
         * @returns {string} The escaped string.
         */
        escapeHtml(str) {
            const div = document.createElement('div');
            div.appendChild(document.createTextNode(str));
            return div.innerHTML;
        }
    }

    // Initialize the PacketCompiler once the DOM is fully loaded
    document.addEventListener('DOMContentLoaded', () => {
        console.log('DOM Content Loaded. Initializing PacketCompiler.');
        new PacketCompiler();
    });
})();
