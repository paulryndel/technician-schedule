document.addEventListener('DOMContentLoaded', () => {
    // --- STATE ---
    let currentDate = new Date();
    let events = {};
    let globalJsonData = [];
    let processedGlobalData = []; // Holds data with corrected dates
    let technicianColors = {};
    let technicianPhotos = {}; // To store photo URLs/data
    let activeFilters = { Technician: new Set(), 'Assigned By': new Set(), Type: new Set(), Status: new Set() };
    const excludedFromAvailable = ['Danuporn', 'Disorn', 'Tossapol']; // Technicians to exclude from 'Available' in calendar view
    let ganttMode = 'technician';
    let chartInstances = {}; // To hold chart instances for destruction
    const P_SCORE_MEANINGS = {
        P1: "Work Performance / Accomplishments",
        P2: "Time Management",
        P3: "Coordination",
        P4: "Learning / Problem Solving",
        P5: "Reporting / Feedback",
        P6: "Discipline"
    };
    const inspirationalQuotes = [
        "The only way to do great work is to love what you do. - Steve Jobs",
        "Success is not the key to happiness. Happiness is the key to success. If you love what you are doing, you will be successful. - Albert Schweitzer",
        "Choose a job you love, and you will never have to work a day in your life. - Confucius",
        "The future depends on what you do today. - Mahatma Gandhi",
        "Believe you can and you're halfway there. - Theodore Roosevelt",
        "The harder I work, the luckier I get. - Samuel Goldwyn",
        "Don't watch the clock; do what it does. Keep going. - Sam Levenson",
        "The only place where success comes before work is in the dictionary. - Vidal Sassoon",
        "Quality is not an act, it is a habit. - Aristotle",
        "The expert in anything was once a beginner.",
        "Strive not to be a success, but rather to be of value. - Albert Einstein",
        "I am a great believer in luck, and I find the harder I work the more I have of it. - Thomas Jefferson",
        "The secret of getting ahead is getting started. - Mark Twain",
        "It’s not what you achieve, it’s what you overcome. That’s what defines your career. - Carlton Fisk",
        "Do the best you can until you know better. Then when you know better, do better. - Maya Angelou",
        "A dream does not become reality through magic; it takes sweat, determination, and hard work. - Colin Powell",
        "The successful warrior is the average man, with laser-like focus. - Bruce Lee",
        "Don't be afraid to give up the good to go for the great. - John D. Rockefeller",
        "Opportunities don't happen. You create them. - Chris Grosser",
        "Success is the sum of small efforts, repeated day-in and day-out. - Robert Collier"
    ];

    // --- DOM ELEMENTS ---
    const monthYearEl = document.getElementById('month-year');
    const calendarContainerEl = document.getElementById('calendar-container');
    const prevMonthBtn = document.getElementById('prev-month');
    const nextMonthBtn = document.getElementById('next-month');
    const fileUploadInput = document.getElementById('file-upload');
    const fileNameEl = document.getElementById('file-name');
    const modal = document.getElementById('event-modal');
    const modalDateEl = document.getElementById('modal-date');
    const modalEventsEl = document.getElementById('modal-events');
    const closeModalBtn = document.getElementById('close-modal');
    const jobDetailModal = document.getElementById('job-detail-modal');
    const jobDetailContent = document.getElementById('job-detail-content');
    const closeJobDetailModalBtn = document.getElementById('close-job-detail-modal');
    const legendDiv = document.getElementById('technician-legend');
    const legendContainer = document.getElementById('legend-container');
    const briefingSection = document.getElementById('briefing-section');
    const filterContainer = document.getElementById('filter-container');
    const technicianFiltersDiv = document.getElementById('technician-filters');
    const assignedByFiltersDiv = document.getElementById('assigned-by-filters');
    const typeFiltersDiv = document.getElementById('type-filters');
    const statusFiltersDiv = document.getElementById('status-filters');
    const clearFiltersBtn = document.getElementById('clear-filters-btn');
    const summarySection = document.getElementById('summary-section');
    const summaryTableContainer = document.getElementById('summary-table-container');
    const summaryMonthInput = document.getElementById('summary-month-input');
    const miniCalendarLegend = document.getElementById('mini-calendar-legend');
    const printSummaryBtn = document.getElementById('print-summary-btn');
    const calendarTab = document.getElementById('calendar-tab');
    const ganttTab = document.getElementById('gantt-tab');
    const reportTab = document.getElementById('report-tab');
    const calendarView = document.getElementById('calendar-view');
    const ganttView = document.getElementById('gantt-view');
    const reportView = document.getElementById('report-view');
    const ganttChartContainer = document.getElementById('gantt-chart-container');
    const ganttMonthYearEl = document.getElementById('gantt-month-year');
    const ganttStartDateInput = document.getElementById('gantt-start-date');
    const ganttEndDateInput = document.getElementById('gantt-end-date');
    const ganttViewModeRadios = document.querySelectorAll('input[name="gantt-mode"]');
    const ganttSortSelect = document.getElementById('gantt-sort');
    const ganttLegend = document.getElementById('gantt-legend');
    const countryTechSelect = document.getElementById('country-tech-select');
    const technicianSkillsContainer = document.getElementById('technician-skills-container');
    const printCalendarBtn = document.getElementById('print-calendar-btn');
    const printGanttBtn = document.getElementById('print-gantt-btn');
    const printReportBtn = document.getElementById('print-report-btn');
    const pscoreLegendEl = document.getElementById('pscore-legend');
    const reportStartDateInput = document.getElementById('report-start-date');
    const reportEndDateInput = document.getElementById('report-end-date');
    const dateDisplayEl = document.getElementById('date-display');
    const timeDisplayEl = document.getElementById('time-display');
    const quoteEl = document.getElementById('inspirational-quote');


    const colorPalette = ['#ef4444', '#f97316', '#eab308', '#84cc16', '#22c55e', '#10b981', '#06b6d4', '#3b82f6', '#6366f1', '#a855f7', '#d946ef', '#ec4899'];
    const specialTechnicianColors = {
        'Somyod': '#8b5cf6' // Violet
    };
    const statusOrder = ['Working', 'On Plan', 'Holding', 'Booking', 'Completed', 'Available'];
    // MODIFIED: Removed 'Cleaning' and 'Cancelled' from statusColors
    const statusColors = {
        'Booking':   { bg: '#38bdf8', text: '#ffffff' }, // Sky Blue
        'On Plan':   { bg: '#facc15', text: '#1f2937' }, // Yellow
        'Holding':   { bg: '#4b5563', text: '#ffffff' }, // Dark Grey
        'Working':   { bg: '#ef4444', text: '#ffffff' }, // Red
        'Completed': { bg: '#22c55e', text: '#ffffff' }, // Green
        'Available': { bg: '#d1d5db', text: '#1f2937' }  // Light Grey
    };

    // Helper function to convert Excel date to a JS Date object, avoiding timezone shifts.
    const excelDateToJSDate = (excelSerial) => {
        if (excelSerial === null || excelSerial === undefined) return null;
        
        // Check if it's a number (Excel serial date)
        if (typeof excelSerial === 'number') {
            // The formula for converting Excel serial numbers to JS timestamp.
            // (excelSerial - 25569) converts the Excel date to days since 1970-01-01.
            // Multiplying by 86400000 converts days to milliseconds.
            // This creates a date object in UTC.
            return new Date((excelSerial - 25569) * 86400000);
        }

        // Check if it's already a Date object (from cellDates:true)
        if (excelSerial instanceof Date) {
            // It's already a date object, just return a new one at UTC midnight
            return new Date(Date.UTC(excelSerial.getUTCFullYear(), excelSerial.getUTCMonth(), excelSerial.getUTCDate()));
        }

        // If it's a string, try parsing it
        if (typeof excelSerial === 'string') {
            const d = new Date(excelSerial);
            if (!isNaN(d.getTime())) {
                return new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
            }
        }
        
        return null;
    };

    const setupTechnicianColors = (jsonData) => {
        const technicians = new Set(jsonData.filter(r => r.Technician).map(r => r.Technician));
        technicianColors = { ...specialTechnicianColors }; // Start with special assignments
        let colorIndex = 0;

        Array.from(technicians).sort().forEach(tech => {
            if (!technicianColors[tech]) {
                technicianColors[tech] = colorPalette[colorIndex % colorPalette.length];
                colorIndex++;
            }
        });
        renderLegend();
    };

    const renderLegend = () => {
        legendContainer.innerHTML = '';
        ganttLegend.innerHTML = '';
        if (Object.keys(technicianColors).length > 0) {
            legendDiv.classList.remove('hidden');
            // Main technician legend with photos
            const techniciansToDisplay = Object.keys(technicianColors)
                .filter(tech => !['Danuporn', 'Tossapol', 'Disorn'].includes(tech))
                .sort();
            techniciansToDisplay.forEach(tech => {
                const legendItem = document.createElement('div');
                legendItem.className = 'flex items-center p-1 bg-white rounded-full shadow-sm';
                const photoSrc = technicianPhotos[tech];
                if (photoSrc) {
                    const img = document.createElement('img');
                    img.src = photoSrc;
                    img.className = 'w-6 h-6 rounded-full mr-2 object-cover';
                    legendItem.appendChild(img);
                }
                legendItem.innerHTML += `<span class="tech-color-dot" style="background-color: ${technicianColors[tech]};"></span><span class="text-sm">${tech}</span>`;
                legendContainer.appendChild(legendItem);
            });

            // Gantt chart status legend
            for (const [status, colors] of Object.entries(statusColors)) {
                 const ganttLegendItem = document.createElement('div');
                ganttLegendItem.className = 'flex items-center';
                ganttLegendItem.innerHTML = `<span class="tech-color-dot" style="background-color: ${colors.bg};"></span><span class="text-sm">${status}</span>`;
                ganttLegend.appendChild(ganttLegendItem);
            }
        } else {
            legendDiv.classList.add('hidden');
        }
    };

    const renderCalendar = () => {
        const year = currentDate.getFullYear();
        const month = currentDate.getMonth();
        monthYearEl.textContent = `${currentDate.toLocaleString('default', { month: 'long' })} ${year}`;
        calendarContainerEl.innerHTML = '';

        const daysOfWeek = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
        daysOfWeek.forEach(day => {
            const headerEl = document.createElement('div');
            headerEl.className = 'text-center font-bold bg-yellow-500 rounded-md h-4 flex items-center justify-center text-xs text-white';
            headerEl.textContent = day;
            calendarContainerEl.appendChild(headerEl);
        });

        const firstDayOfMonth = new Date(year, month, 1).getDay();
        const lastDateOfMonth = new Date(year, month + 1, 0).getDate();
        const today = new Date();

        for (let i = 0; i < firstDayOfMonth; i++) calendarContainerEl.appendChild(document.createElement('div'));

        for (let i = 1; i <= lastDateOfMonth; i++) {
            const dayDiv = document.createElement('div');
            dayDiv.className = 'day';
            const dayContent = document.createElement('div');
            dayContent.className = 'day-content';
            const dayNumberDiv = document.createElement('div');
            dayNumberDiv.className = 'day-number';
            dayNumberDiv.textContent = i;
            dayDiv.append(dayNumberDiv, dayContent);
            const dateStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(i).padStart(2, '0')}`;
            if (year === today.getFullYear() && month === today.getMonth() && i === today.getDate()) dayDiv.classList.add('today');

            const currentDay = new Date(year, month, i).getDay();

            if (currentDay === 0) {
                dayDiv.classList.add('sunday-off');
                dayContent.innerHTML = '';
            } else {
                const dayEvents = (events[dateStr] || []).filter(event => {
                    const techMatch = activeFilters.Technician.size === 0 || activeFilters.Technician.has(event.Technician);
                    const assignedByMatch = activeFilters['Assigned By'].size === 0 || activeFilters['Assigned By'].has(event['Assigned By']);
                    const typeMatch = activeFilters.Type.size === 0 || activeFilters.Type.has(event.Type);
                    return techMatch && assignedByMatch && typeMatch;
                });

                if(dayEvents.length > 0) {
                    dayDiv.classList.add('day-with-event');
                    dayDiv.addEventListener('click', () => showEventModal(dateStr));
                }

                // Only show Technician and Type
                dayEvents.forEach(event => {
                    if (event.Technician) {
                        const tag = document.createElement('div');
                        tag.className = 'tech-tag';
                        tag.style.backgroundColor = technicianColors[event.Technician] || '#cccccc';
                        tag.innerHTML = `
                            <span class="tech-tag-name">${event.Technician}</span>
                            <span class="tech-tag-details">${event.Type || ''}</span>
                        `;
                        dayContent.appendChild(tag);
                    }
                });
            }
            calendarContainerEl.appendChild(dayDiv);
        }
        calendarContainerEl.classList.add('calendar-grid-body');
    };

    const renderGanttChart = () => {
        if (processedGlobalData.length === 0) {
            ganttChartContainer.innerHTML = '<p class="text-center p-4">Upload an Excel file to see the Gantt Chart.</p>';
            return;
        }

        const today = new Date();
        const ganttStartDate = new Date(ganttStartDateInput.value + 'T00:00:00Z');
        const ganttEndDate = new Date(ganttEndDateInput.value + 'T23:59:59Z');

        if (isNaN(ganttStartDate) || isNaN(ganttEndDate) || ganttStartDate > ganttEndDate) {
            ganttChartContainer.innerHTML = '<p class="text-center p-4 text-red-500">Invalid date range selected.</p>';
            return;
        }

        ganttMonthYearEl.textContent = `Gantt Chart: ${ganttStartDate.toLocaleDateString('default', {timeZone: 'UTC'})} - ${ganttEndDate.toLocaleDateString('default', {timeZone: 'UTC'})}`;
        ganttChartContainer.innerHTML = '';

        const allDays = [];
        const monthsInRange = {};
        for (let d = new Date(ganttStartDate); d <= ganttEndDate; d.setUTCDate(d.getUTCDate() + 1)) {
            allDays.push(new Date(d));
            const monthKey = `${d.getUTCFullYear()}-${d.getUTCMonth()}`;
            if (!monthsInRange[monthKey]) {
                monthsInRange[monthKey] = { name: d.toLocaleString('default', { month: 'long', year: 'numeric', timeZone: 'UTC'}), days: 0 };
            }
            monthsInRange[monthKey].days++;
        }
        const totalDays = allDays.length;

        const grid = document.createElement('div');
        grid.className = 'gantt-grid';
        
        // --- Header Row ---
        const header = document.createElement('div');
        header.className = 'gantt-header';
        
        const headerLabel = document.createElement('div');
        headerLabel.className = 'gantt-tech-label';
        headerLabel.style.position = 'sticky';
        headerLabel.style.top = '0';
        headerLabel.innerHTML = `<div class="h-[49px] flex items-center">${ganttMode === 'technician' ? 'Technician' : 'Customer'}</div>`; // Match header height
        header.appendChild(headerLabel);

        const timelineHeaderContainer = document.createElement('div');
        timelineHeaderContainer.className = 'gantt-timeline-header-container';
        
        const monthHeader = document.createElement('div');
        monthHeader.className = 'gantt-month-header';
        monthHeader.style.gridTemplateColumns = Object.values(monthsInRange).map(m => `${m.days}fr`).join(' ');
        
        Object.values(monthsInRange).forEach(m => {
            const monthLabel = document.createElement('div');
            monthLabel.className = 'gantt-month-label';
            monthLabel.textContent = m.name;
            monthHeader.appendChild(monthLabel);
        });

        const dayHeader = document.createElement('div');
        dayHeader.className = 'gantt-timeline-header';
        dayHeader.style.gridTemplateColumns = `repeat(${totalDays}, 1fr)`;

        allDays.forEach(d => {
            const dayLabel = document.createElement('div');
            dayLabel.className = 'gantt-day-label';
            dayLabel.textContent = d.getUTCDate();
            if (d.getUTCDay() === 0) dayLabel.classList.add('gantt-sunday');
            dayHeader.appendChild(dayLabel);
        });

        timelineHeaderContainer.append(monthHeader, dayHeader);
        header.appendChild(timelineHeaderContainer);
        grid.appendChild(header);

        // --- Data Rows ---
        const jobsInView = processedGlobalData.filter(job => {
            if (!job.processedStartDate || !job.processedEndDate) return false;
            const jobStart = new Date(job.processedStartDate);
            const jobEnd = new Date(job.processedEndDate);
            if (jobStart > ganttEndDate || jobEnd < ganttStartDate) return false;

            const techMatch = activeFilters.Technician.size === 0 || activeFilters.Technician.has(job.Technician);
            const assignedByMatch = activeFilters['Assigned By'].size === 0 || activeFilters['Assigned By'].has(job['Assigned By']);
            const typeMatch = activeFilters.Type.size === 0 || activeFilters.Type.has(job.Type);
            const status = job.Status || 'On Plan';
            const statusMatch = activeFilters.Status.size === 0 || activeFilters.Status.has(status);

            return techMatch && assignedByMatch && typeMatch && statusMatch;
        });

        let itemsToDisplay;
        if (ganttMode === 'technician') {
            itemsToDisplay = Object.keys(technicianColors)
                .filter(tech => !['Danuporn', 'Tossapol', 'Disorn'].includes(tech))
                .sort();
        } else {
            const customersInView = new Set(jobsInView.filter(r => r.Customer).map(r => r.Customer));
            let customerArray = Array.from(customersInView);
            const sortMode = ganttSortSelect.value;
            if (sortMode === 'date') {
                const customerStartDateMap = new Map();
                customerArray.forEach(customer => {
                    const firstJob = jobsInView
                        .filter(job => job.Customer === customer)
                        .sort((a,b) => a.processedStartDate - b.processedStartDate)[0];
                    customerStartDateMap.set(customer, firstJob.processedStartDate);
                });
                customerArray.sort((a,b) => {
                    const dateA = customerStartDateMap.get(a);
                    const dateB = customerStartDateMap.get(b);
                    return dateA - dateB;
                });
            } else { // 'name'
                customerArray.sort();
            }
            itemsToDisplay = customerArray;
        }

        itemsToDisplay.forEach(item => {
            const row = document.createElement('div');
            row.className = 'gantt-row';

            const itemLabel = document.createElement('div');
            itemLabel.className = 'gantt-tech-label';
            
            if (ganttMode === 'technician') {
                const photoSrc = technicianPhotos[item];
                if (photoSrc) {
                    const img = document.createElement('img');
                    img.src = photoSrc;
                    img.alt = item;
                    img.className = 'gantt-tech-photo';
                    img.onerror = function() { this.style.display = 'none'; };
                    itemLabel.appendChild(img);
                }
            }
            const nameSpan = document.createElement('span');
            nameSpan.textContent = item;
            itemLabel.appendChild(nameSpan);
            row.appendChild(itemLabel);

            const timeline = document.createElement('div');
            timeline.className = 'gantt-timeline';
            timeline.style.gridTemplateColumns = `repeat(${totalDays}, 1fr)`;

            allDays.forEach((d, index) => {
                if (d.getUTCDay() === 0) {
                    const sundayMarker = document.createElement('div');
                    sundayMarker.className = 'gantt-sunday-marker';
                    sundayMarker.style.width = `${(1 / totalDays) * 100}%`;
                    sundayMarker.style.left = `${(index / totalDays) * 100}%`;
                    timeline.appendChild(sundayMarker);
                }
            });

            const todayTime = new Date(today.getFullYear(), today.getMonth(), today.getDate()).getTime();
            if (todayTime >= ganttStartDate.getTime() && todayTime <= ganttEndDate.getTime()) {
                const todayMarker = document.createElement('div');
                todayMarker.className = 'gantt-today-marker';
                const dayOffset = (todayTime - ganttStartDate.getTime()) / (1000 * 3600 * 24);
                todayMarker.style.left = `calc(${(dayOffset / totalDays) * 100}% + (${(1 / totalDays) * 100 / 2}% - 1px))`;
                timeline.appendChild(todayMarker);
            }

            const jobsForItem = jobsInView.filter(job => {
                const key = ganttMode === 'technician' ? job.Technician : job.Customer;
                return key === item;
            }).sort((a, b) => a.processedStartDate - b.processedStartDate);

            const lanes = [];
            jobsForItem.forEach(job => {
                let placed = false;
                for (const lane of lanes) {
                    const lastJobInLane = lane[lane.length - 1];
                    if (job.processedStartDate > lastJobInLane.processedEndDate) {
                        lane.push(job);
                        placed = true;
                        break;
                    }
                }
                if (!placed) {
                    lanes.push([job]);
                }
            });

            const numLanes = lanes.length || 1;
            lanes.forEach((lane, laneIndex) => {
                lane.forEach(job => {
                    let jobStart = new Date(job.processedStartDate);
                    let jobEnd = new Date(job.processedEndDate);
                    jobStart = new Date(Math.max(jobStart, ganttStartDate));
                    jobEnd = new Date(Math.min(jobEnd, ganttEndDate));
                    
                    const startOffset = (jobStart.getTime() - ganttStartDate.getTime()) / (1000 * 3600 * 24);
                    const duration = (jobEnd.getTime() - jobStart.getTime()) / (1000 * 3600 * 24) + 1;
                    
                    const bar = document.createElement('div');
                    bar.className = 'gantt-bar';
                    const barHeight = 100 / numLanes;
                    bar.style.height = `calc(${barHeight}% - 4px)`;
                    bar.style.top = `calc(${laneIndex * barHeight}% + 2px)`;
                    bar.style.left = `${(startOffset / totalDays) * 100}%`;
                    bar.style.width = `calc(${(duration / totalDays) * 100}% - 2px)`;
                    const status = job.Status || 'On Plan';
                    const statusColorInfo = statusColors[status] || {bg: '#cccccc', text: '#000000'};
                    bar.style.backgroundColor = statusColorInfo.bg;
                    bar.style.color = statusColorInfo.text;
                    bar.style.zIndex = '10';
                    
                    if (ganttMode === 'customer') {
                        const photoSrc = technicianPhotos[job.Technician];
                        if (photoSrc) {
                            const img = document.createElement('img');
                            img.src = photoSrc;
                            img.className = 'gantt-bar-photo';
                            bar.appendChild(img);
                        }
                        const textSpan = document.createElement('span');
                        textSpan.textContent = job.Technician;
                        bar.appendChild(textSpan);
                    } else {
                        const textSpan = document.createElement('span');
                        textSpan.textContent = job.Customer || job.Type;
                        bar.appendChild(textSpan);
                    }

                    bar.title = `${job.Technician} - ${job.Customer} (${job.Type})\nDuration: ${duration} day(s)\nAssigned by: ${job['Assigned By'] || 'N/A'}\nStatus: ${status}`;
                    bar.addEventListener('click', () => showJobDetailModal(job));
                    
                    timeline.appendChild(bar);
                });
            });

            row.appendChild(timeline);
            grid.appendChild(row);
        });

        ganttChartContainer.appendChild(grid);
    };
    

    const showEventModal = (dateStr) => {
        const dateEvents = events[dateStr];
        if (!dateEvents) return;

        const displayDate = new Date(dateStr + 'T12:00:00Z'); // Use noon UTC to ensure correct date display
        modalDateEl.textContent = displayDate.toLocaleDateString('default', { timeZone: 'UTC', weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });
        modalEventsEl.innerHTML = '';
        briefingSection.innerHTML = '';
        
        const workingEntries = dateEvents.filter(e => e.Technician);
        const otherEntries = dateEvents.filter(e => !e.Technician);

        if (workingEntries.length > 0) {
            const workingDiv = document.createElement('div');
            workingDiv.innerHTML = '<h4 class="text-md font-semibold text-gray-800 mb-2 border-b pb-1">Technician Assignments</h4>';
            const list = document.createElement('div');
            list.className = 'space-y-3';
            workingEntries.forEach(event => {
                const eventDiv = document.createElement('div');
                const techColor = technicianColors[event.Technician] || '#e5e7eb';
                eventDiv.className = 'p-3 rounded-lg';
                eventDiv.style.borderLeft = `5px solid ${techColor}`;
                eventDiv.style.backgroundColor = `${techColor}1A`;
                let detailsHtml = `<p class="font-bold" style="color: ${techColor};">${event.Technician}</p><p class="text-sm text-gray-600">${event.Description || ''}</p><div class="mt-2 text-xs text-gray-500 grid grid-cols-2 gap-x-4">`;
                for (const [key, value] of Object.entries(event)) {
                    if (value && !['Technician', 'Description', 'Plan Start', 'Plan Finish', 'processedStartDate', 'processedEndDate'].includes(key)) detailsHtml += `<div><strong class="font-semibold">${key}:</strong> ${value}</div>`;
                }
                detailsHtml += '</div>';
                eventDiv.innerHTML = detailsHtml;
                list.appendChild(eventDiv);
            });
            workingDiv.appendChild(list);
            modalEventsEl.appendChild(workingDiv);
        }

        if (otherEntries.length > 0) {
            const otherDiv = document.createElement('div');
            otherDiv.innerHTML = '<h4 class="text-md font-semibold text-gray-800 mt-4 mb-2 border-b pb-1">Other Scheduled Items</h4>';
            const list = document.createElement('div');
            list.className = 'space-y-3';
            otherEntries.forEach(event => {
                const eventDiv = document.createElement('div');
                eventDiv.className = 'p-3 bg-blue-50 rounded-lg border-l-4 border-blue-300';
                let detailsHtml = `<p class="font-bold text-blue-800">${event.Description || event.Type || 'Scheduled Item'}</p><div class="mt-2 text-xs text-gray-500 grid grid-cols-2 gap-x-4">`;
                for (const [key, value] of Object.entries(event)) {
                     if (value && !['Description', 'Plan Start', 'Plan Finish', 'processedStartDate', 'processedEndDate'].includes(key)) detailsHtml += `<div><strong class="font-semibold">${key}:</strong> ${value}</div>`;
                }
                detailsHtml += '</div>';
                eventDiv.innerHTML = detailsHtml;
                list.appendChild(eventDiv);
            });
            otherDiv.appendChild(list);
            modalEventsEl.appendChild(otherDiv);
        }
        modal.classList.remove('hidden');
    };
    
    const showJobDetailModal = (job) => {
        const modalContent = jobDetailModal.querySelector('.modal-content');
        
        // Remove any existing photo
        const existingPhoto = modalContent.querySelector('.job-detail-photo');
        if (existingPhoto) {
            existingPhoto.remove();
        }

        jobDetailContent.innerHTML = '';
        
        // Add technician photo to the top right
        const photoSrc = technicianPhotos[job.Technician];
        if (photoSrc) {
            const photoContainer = document.createElement('div');
            photoContainer.className = 'job-detail-photo absolute top-6 right-6';
            const img = document.createElement('img');
            img.src = photoSrc;
            img.alt = job.Technician;
            img.className = 'w-20 h-20 rounded-full object-cover border-4 border-white shadow-lg';
            photoContainer.appendChild(img);
            modalContent.appendChild(photoContainer);
        }

        const displayOrder = ['Customer', 'Technician', 'Type', 'Description', 'Status', 'Plan Start', 'Plan Finish', 'Assigned By'];
        
        displayOrder.forEach(key => {
            if (job[key]) {
                const detailRow = document.createElement('div');
                detailRow.className = 'grid grid-cols-3 gap-4 py-2 border-b border-yellow-200';
                const keyEl = document.createElement('strong');
                keyEl.className = 'col-span-1 text-yellow-800 font-semibold';
                keyEl.textContent = key;
                const valueEl = document.createElement('span');
                valueEl.className = 'col-span-2 text-gray-800';
                let value = job[key];
                if (value instanceof Date) {
                    value = value.toLocaleDateString('default', { timeZone: 'UTC', year: 'numeric', month: 'long', day: 'numeric' });
                }
                valueEl.textContent = value;
                detailRow.append(keyEl, valueEl);
                jobDetailContent.appendChild(detailRow);
            }
        });

        jobDetailModal.classList.remove('hidden');
    };

    const hideModal = () => modal.classList.add('hidden');
    const hideJobDetailModal = () => {
        const modalContent = jobDetailModal.querySelector('.modal-content');
        const existingPhoto = modalContent.querySelector('.job-detail-photo');
        if (existingPhoto) {
            existingPhoto.remove();
        }
        jobDetailModal.classList.add('hidden');
    }
    
    const renderFilters = (jsonData) => {
        filterContainer.classList.remove('hidden');
        
        const createFilterDropdown = (items, type) => {
            const container = document.getElementById(`${type.toLowerCase().replace(' ', '-')}-filters`);
            container.innerHTML = '';

            const button = document.createElement('button');
            button.className = 'filter-button';
            button.innerHTML = `${type} <span id="${type.toLowerCase().replace(' ', '-')}-badge" class="hidden filter-badge">0</span>`;
            
            const dropdown = document.createElement('div');
            dropdown.className = 'filter-dropdown hidden';
            
            items.forEach(item => {
                const label = document.createElement('label');
                label.className = 'flex items-center space-x-2 p-1 hover:bg-gray-100 rounded cursor-pointer';
                const checkbox = document.createElement('input');
                checkbox.type = 'checkbox';
                checkbox.className = 'form-checkbox h-4 w-4 text-blue-600 rounded';
                checkbox.dataset.filterType = type;
                checkbox.dataset.filterValue = item;
                label.append(checkbox, document.createTextNode(item));
                dropdown.appendChild(label);
            });

            container.append(button, dropdown);

            button.addEventListener('click', (e) => {
                e.stopPropagation();
                // Hide other dropdowns
                document.querySelectorAll('.filter-dropdown').forEach(d => {
                    if (d !== dropdown) d.classList.add('hidden');
                });
                dropdown.classList.toggle('hidden');
            });
        };
        
        const technicians = [...new Set(jsonData.filter(r => r.Technician).map(r => r.Technician))].sort();
        const assignedBy = [...new Set(jsonData.filter(r => r['Assigned By']).map(r => r['Assigned By']))].sort();
        const types = [...new Set(jsonData.filter(r => r.Type).map(r => r.Type))].sort();

        createFilterDropdown(technicians, 'Technician');
        createFilterDropdown(assignedBy, 'Assigned By');
        createFilterDropdown(types, 'Type');
        createFilterDropdown(statusOrder, 'Status');
    };

    const updateFilterBadges = () => {
        for (const type in activeFilters) {
            const badge = document.getElementById(`${type.toLowerCase().replace(' ', '-')}-badge`);
            if (badge) {
                const count = activeFilters[type].size;
                if (count > 0) {
                    badge.textContent = count;
                    badge.classList.remove('hidden');
                } else {
                    badge.classList.add('hidden');
                }
            }
        }
    };

    const createMiniCalendar = (scheduledDates, year, month, dailyJobsMap) => {
        const miniCalContainer = document.createElement('div');
        miniCalContainer.className = 'mini-calendar';

        const header = '<div>S</div><div>M</div><div>T</div><div>W</div><div>T</div><div>F</div><div>S</div>';
        miniCalContainer.innerHTML = `<div class="mini-calendar-header">${header}</div>`;
        
        const body = document.createElement('div');
        body.className = 'mini-calendar-body';

        const firstDay = new Date(year, month, 1).getDay();
        const daysInMonth = new Date(year, month + 1, 0).getDate();

        for (let i = 0; i < firstDay; i++) {
            body.innerHTML += '<div></div>';
        }

        for (let day = 1; day <= daysInMonth; day++) {
            const dayCell = document.createElement('div');
            dayCell.textContent = day;
            
            const currentDateObj = new Date(Date.UTC(year, month, day));
            const isSunday = currentDateObj.getUTCDay() === 0;
            const isScheduled = scheduledDates.has(currentDateObj.toDateString());

            let classes = ['mini-calendar-day'];

            if (isSunday) {
                classes.push('mini-calendar-sunday');
            } else if (isScheduled) {
                const jobsForDay = dailyJobsMap.get(day);
                const uniqueStatuses = [...new Set(jobsForDay.map(job => job.status || 'On Plan'))];
                const numStatuses = uniqueStatuses.length;

                if (numStatuses === 1) {
                    const colorInfo = statusColors[uniqueStatuses[0]] || statusColors['Available'];
                    dayCell.style.backgroundColor = colorInfo.bg;
                    dayCell.style.color = colorInfo.text || 'black';
                } else if (numStatuses > 1) {
                    const colorStops = uniqueStatuses.map((status, index) => {
                        const color = (statusColors[status] || statusColors['Available']).bg;
                        const startPercent = (100 / numStatuses) * index;
                        const endPercent = (100 / numStatuses) * (index + 1);
                        return `${color} ${startPercent}% ${endPercent}%`;
                    }).join(', ');
                    
                    dayCell.style.background = `linear-gradient(to right, ${colorStops})`;
                    dayCell.style.color = 'white';
                    dayCell.style.textShadow = '1px 1px 1px rgba(0,0,0,0.4)';
                } else {
                    const colorInfo = statusColors['Available'];
                    dayCell.style.backgroundColor = colorInfo.bg;
                    dayCell.style.color = colorInfo.text || 'black';
                }
            } else {
                // Not Sunday and not scheduled, so it's 'Available'
                const colorInfo = statusColors['Available'];
                dayCell.style.backgroundColor = colorInfo.bg;
                dayCell.style.color = colorInfo.text || 'black';
            }
            
            dayCell.className = classes.join(' ');
            body.appendChild(dayCell);
        }

        miniCalContainer.appendChild(body);
        return miniCalContainer;
    };
    
    const renderSummaryTable = (date) => {
        const targetDate = date || currentDate;
        if (processedGlobalData.length === 0) {
            summarySection.classList.add('hidden');
            return;
        }
        summarySection.classList.remove('hidden');
        
        summaryTableContainer.innerHTML = ''; // Clear previous table
        
        miniCalendarLegend.innerHTML = '';
        // MODIFIED: Added check to exclude certain statuses from legend
        Object.entries(statusColors).forEach(([status, colors]) => {
            if (status === 'Cancelled' || status === 'Cleaning') return;
            const legendItem = document.createElement('div');
            legendItem.className = 'flex items-center';
            const colorBox = document.createElement('span');
            colorBox.className = 'w-3 h-3 rounded-sm mr-1';
            colorBox.style.backgroundColor = colors.bg;
            colorBox.style.border = '1px solid black'; // Added black border
            legendItem.append(colorBox, document.createTextNode(status));
            miniCalendarLegend.appendChild(legendItem);
        });

        const table = document.createElement('div');
        table.className = 'summary-table';

        ['Technician', ...statusOrder].forEach(headerText => {
            if (headerText === 'Cancelled' || headerText === 'Cleaning') return;

            const headerCell = document.createElement('div');
            headerCell.className = 'summary-header';
            headerCell.textContent = headerText;
            if(statusColors[headerText]) {
                headerCell.style.backgroundColor = statusColors[headerText].bg;
                headerCell.style.color = statusColors[headerText].text;
            }
            table.appendChild(headerCell);
        });
        
        const techniciansToDisplay = Object.keys(technicianColors).filter(tech => !excludedFromAvailable.includes(tech)).sort();

        techniciansToDisplay.forEach(tech => {
            const techCell = document.createElement('div');
            techCell.className = 'summary-cell summary-cell-tech flex flex-col items-center justify-center p-2';
            const photoSrc = technicianPhotos[tech];
            const techColor = technicianColors[tech] || '#f3f4f6';
            
            const photoContainer = document.createElement('div');
            photoContainer.className = 'w-20 h-20 rounded-lg mb-2 border-2 border-white shadow-lg flex items-center justify-center p-1';
            photoContainer.style.backgroundColor = techColor;

            if (photoSrc) {
                const img = document.createElement('img');
                img.src = photoSrc;
                img.alt = tech;
                img.className = 'w-full h-full rounded-md object-contain';
                const fallbackIcon = document.createElement('div');
                fallbackIcon.className = 'w-full h-full hidden items-center justify-center';
                fallbackIcon.innerHTML = `<svg class="w-10 h-10 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M16 7a4 4 0 11-8 0 4 4 0 018 0zM12 14a7 7 0 00-7 7h14a7 7 0 00-7-7z"></path></svg>`;
                img.onerror = function() {
                    this.style.display = 'none';
                    fallbackIcon.style.display = 'flex';
                };
                photoContainer.appendChild(img);
                photoContainer.appendChild(fallbackIcon);
            } else {
                photoContainer.innerHTML = `<svg class="w-10 h-10 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M16 7a4 4 0 11-8 0 4 4 0 018 0zM12 14a7 7 0 00-7 7h14a7 7 0 00-7-7z"></path></svg>`;
            }
            
            const techNameSpan = document.createElement('span');
            techNameSpan.className = 'text-center mt-1';
            techNameSpan.textContent = tech;

            techCell.appendChild(photoContainer);
            techCell.appendChild(techNameSpan);
            table.appendChild(techCell);

            const monthStart = new Date(Date.UTC(targetDate.getFullYear(), targetDate.getMonth(), 1));
            const monthEnd = new Date(Date.UTC(targetDate.getFullYear(), targetDate.getMonth() + 1, 0));
            
            const techJobsInMonth = processedGlobalData.filter(job => {
                if (job.Technician !== tech) return false;
                const startDate = job.processedStartDate;
                if (!startDate) return false;
                const endDate = job.processedEndDate || startDate;
                return startDate <= monthEnd && endDate >= monthStart;
            });

            const jobsByStatus = {};
            statusOrder.forEach(s => jobsByStatus[s] = []);
            techJobsInMonth.forEach(job => {
                const status = job.Status || 'On Plan';
                if (jobsByStatus[status]) {
                    jobsByStatus[status].push(job);
                }
            });

            let scheduledDates = new Set();
            let dailyJobsMap = new Map();
            techJobsInMonth.forEach(job => {
                let startDate = job.processedStartDate;
                if (startDate) {
                    let endDate = job.processedEndDate || new Date(startDate);
                    for (let d = new Date(startDate); d <= endDate; d.setUTCDate(d.getUTCDate() + 1)) {
                        if (d.getUTCFullYear() === targetDate.getFullYear() && d.getUTCMonth() === targetDate.getMonth()) {
                            const day = d.getUTCDate();
                            scheduledDates.add(new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), day)).toDateString());
                            if (!dailyJobsMap.has(day)) dailyJobsMap.set(day, []);
                            dailyJobsMap.get(day).push({
                                customer: job.Customer,
                                type: job.Type,
                                status: job.Status || 'On Plan'
                            });
                        }
                    }
                }
            });

            statusOrder.forEach(status => {
                if (status === 'Cancelled' || status === 'Cleaning') return;

                const statusCell = document.createElement('div');
                statusCell.className = 'summary-cell';
                if (status !== 'Available') {
                    const jobsToDisplay = jobsByStatus[status];
                    if (jobsToDisplay) {
                        jobsToDisplay.sort((a, b) => (a.processedStartDate || 0) - (b.processedStartDate || 0));

                        jobsToDisplay.forEach((job, index) => {
                            const jobDiv = document.createElement('div');
                            jobDiv.className = 'summary-job text-left p-1 flex items-start';

                            const workId = job['Work ID'] || 'N/A';
                            const customer = job.Customer || 'N/A';
                            const type = job.Type || 'N/A';
                            
                            const startDate = job.processedStartDate;
                            const endDate = job.processedEndDate || startDate;

                            let dateString;
                            if (startDate) {
                                const startMonth = startDate.toLocaleString('default', { month: 'short', timeZone: 'UTC' });
                                const endMonth = endDate.toLocaleString('default', { month: 'short', timeZone: 'UTC' });

                                if (startDate.getTime() !== endDate.getTime()) {
                                    if (startMonth === endMonth) {
                                        dateString = `${startDate.getUTCDate()}-${endDate.getUTCDate()} ${endMonth}`;
                                    } else {
                                        dateString = `${startDate.getUTCDate()} ${startMonth} - ${endDate.getUTCDate()} ${endMonth}`;
                                    }
                                } else {
                                    dateString = `${startDate.getUTCDate()} ${startMonth}`;
                                }
                            } else {
                                dateString = 'N/A';
                            }
                            
                            jobDiv.innerHTML = `
                                <span class="w-6 shrink-0 text-right pr-2">${index + 1}.</span>
                                <div>
                                    <b>${workId}</b>  :  <span class="font-semibold">${customer}</span> | ${dateString} | ${type}
                                </div>
                            `;
                            statusCell.appendChild(jobDiv);
                        });
                    }
                } else {
                    const miniCal = createMiniCalendar(scheduledDates, targetDate.getFullYear(), targetDate.getMonth(), dailyJobsMap);
                    statusCell.appendChild(miniCal);
                }
                table.appendChild(statusCell);
            });
        });

        summaryTableContainer.appendChild(table);
    };

    const handleFileUpload = (e) => {
        const file = e.target.files[0];
        if (!file) return;
        fileNameEl.textContent = `Processing: ${file.name}...`;
        const reader = new FileReader();
        reader.onload = (event) => {
            try {
                const data = new Uint8Array(event.target.result);
                const workbook = XLSX.read(data, { type: 'array', cellDates: true });
                
                const worklistSheetName = 'Work List';
                const worksheet = workbook.Sheets[worklistSheetName];
                if (!worksheet) throw new Error(`Sheet "${worklistSheetName}" not found in the Excel file.`);
                globalJsonData = XLSX.utils.sheet_to_json(worksheet, { range: 1 });

                const infoSheetName = 'Information';
                const infoWorksheet = workbook.Sheets[infoSheetName];
                technicianPhotos = {};

                if (infoWorksheet) {
                    const infoData = XLSX.utils.sheet_to_json(infoWorksheet);
                    console.log("Found 'Information' sheet. Data:", infoData);
                    let photoCount = 0;
                    infoData.forEach(row => {
                        const nameKey = Object.keys(row).find(k => k.toLowerCase() === 'name');
                        const photoKey = Object.keys(row).find(k => k.toLowerCase() === 'photo');

                        if (nameKey && photoKey && row[nameKey] && row[photoKey]) {
                            const name = String(row[nameKey]).trim();
                            const photoUrl = String(row[photoKey]);
                            technicianPhotos[name] = photoUrl;
                            photoCount++;
                        }
                    });
                    if (photoCount > 0) {
                        console.log(`Successfully loaded ${photoCount} technician photos.`, technicianPhotos);
                    } else {
                        console.warn("No technician photos were loaded. Check the 'Information' sheet. It must contain 'Name' and 'Photo' columns. The 'Photo' column should contain image URLs.");
                    }
                } else {
                    console.warn(`Sheet "${infoSheetName}" not found. No photos will be displayed.`);
                }

                processedGlobalData = globalJsonData.map(row => {
                    const newRow = {...row};
                    // ADD ONE DAY TO DATES
                    let startDate = excelDateToJSDate(row['Plan Start']);
                    if (startDate) {
                        startDate.setUTCDate(startDate.getUTCDate() + 1);
                    }
                    newRow.processedStartDate = startDate;

                    let endDate = excelDateToJSDate(row['Plan Finish']);
                    if (endDate) {
                        endDate.setUTCDate(endDate.getUTCDate() + 1);
                    }
                    if (!endDate && startDate) {
                        endDate = new Date(startDate);
                    }
                    newRow.processedEndDate = endDate;
                    return newRow;
                });

                setupTechnicianColors(processedGlobalData);
                const newEvents = {};

                processedGlobalData.forEach(row => {
                    let startDate = row.processedStartDate;
                    if (startDate) {
                        let endDate = row.processedEndDate || new Date(startDate);
                        for (let d = new Date(startDate); d <= endDate; d.setUTCDate(d.getUTCDate() + 1)) {
                            const dateStr = `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}-${String(d.getUTCDate()).padStart(2, '0')}`;
                            if (!newEvents[dateStr]) newEvents[dateStr] = [];
                            newEvents[dateStr].push(row);
                        }
                    }
                });
                events = newEvents;
                fileNameEl.textContent = `Successfully loaded: ${file.name}`;
                renderFilters(processedGlobalData);
                updateViews();
            } catch (error) {
                console.error("Error processing Excel file:", error);
                fileNameEl.textContent = `Error: ${error.message}`;
            }
        };
        reader.onerror = () => fileNameEl.textContent = "Error reading file.";
        reader.readAsArrayBuffer(file);
    };
    
    function getPreconfiguredSummary(scores) {
        const avg = Object.values(scores).reduce((a, b) => a + b, 0) / Object.values(scores).length;
        let summary = "";
        let recommendation = "";

        if (avg >= 9.5) {
            summary = "Exhibits outstanding performance across all metrics, serving as a model for the team.";
            recommendation = "Continue to mentor other technicians and lead by example. Explore opportunities for advanced projects or leadership roles.";
        } else if (avg >= 8.5) {
            summary = "Demonstrates excellent proficiency and consistency in all key areas of work.";
            recommendation = "Focus on sharing best practices with peers to help elevate the team's overall performance.";
        } else if (avg >= 7.5) {
            summary = "Shows strong, reliable performance with solid results in most areas.";
            recommendation = "Identify the lowest scoring area and seek targeted training or mentorship to achieve excellence across the board.";
        } else if (avg >= 6.5) {
            summary = "Maintains a good, consistent level of performance meeting all job expectations.";
            recommendation = "Proactively seek feedback on complex tasks to further refine problem-solving and reporting skills.";
        } else if (avg >= 5.0) {
            summary = "Performance is satisfactory and meets the required standards, with some room for growth.";
            recommendation = "Concentrate on improving time management and coordination to increase overall efficiency and output.";
        } else if (avg >= 4.0) {
            summary = "Performance is inconsistent, meeting basic requirements but showing clear areas for improvement.";
            recommendation = "A structured development plan focusing on discipline and reporting is recommended to build consistency.";
        } else {
            summary = "Shows significant challenges in meeting performance expectations across multiple areas.";
            recommendation = "Immediate intervention and a foundational training plan are crucial to address core skill gaps.";
        }
        
        let lowestKey = 'P1';
        let highestKey = 'P1';
        if (Object.keys(scores).length > 0 && Object.values(scores).some(s => s > 0)) {
            for (const key in scores) {
                if (scores[key] < scores[lowestKey]) lowestKey = key;
                if (scores[key] > scores[highestKey]) highestKey = key;
            }
            summary = `This technician's greatest strength appears to be in ${P_SCORE_MEANINGS[highestKey]}. ${summary}`;
            recommendation = `To improve, a focus on ${P_SCORE_MEANINGS[lowestKey]} is suggested. ${recommendation}`;
        }


        return `<strong>Summary:</strong> ${summary}<br><br><strong>Recommendation:</strong> ${recommendation}`;
    }

    function renderReportCharts() {
        Chart.register(ChartDataLabels);

        for(const key in chartInstances) {
            if(chartInstances[key]) chartInstances[key].destroy();
        }
        chartInstances = {};
        
        technicianSkillsContainer.innerHTML = '';
        pscoreLegendEl.innerHTML = Object.entries(P_SCORE_MEANINGS)
            .map(([key, value]) => `<li><strong class="text-yellow-800">${key}:</strong> ${value}</li>`)
            .join('');

        if (!processedGlobalData || processedGlobalData.length === 0) {
            reportView.querySelector('.grid').innerHTML = '<p class="text-center col-span-full text-red-700">No data loaded. Please upload an Excel file.</p>';
            return;
        }

        const startDate = reportStartDateInput.value ? new Date(reportStartDateInput.value) : null;
        const endDate = reportEndDateInput.value ? new Date(reportEndDateInput.value) : null;

        const filteredReportData = processedGlobalData.filter(job => {
            if (!job.processedStartDate) return false;
            const jobDate = new Date(job.processedStartDate);
            if (startDate && jobDate < startDate) return false;
            if (endDate && jobDate > endDate) return false;
            return true;
        });
        
        // MODIFIED: Added 'Tossapol' to the exclusion list for reports
        const reportTechsToExclude = ['Danuporn', 'Disorn', 'Tossapol'];
        const allTechs = Array.from(new Set(filteredReportData.map(r => r.Technician).filter(Boolean)))
            .filter(tech => !reportTechsToExclude.includes(tech))
            .sort();

        // 1. Job Duration by Grade
        const techGradeDays = {};
        const allGrades = ['A','B','C','D'];
        filteredReportData.forEach(row => {
            const tech = row.Technician;
            if (!allTechs.includes(tech)) return;
            const grade = (row.Grade || '').toString().toUpperCase();
            if (!tech || !allGrades.includes(grade)) return;
            
            const start = row.processedStartDate;
            const end = row.processedEndDate || start;
            let days = 0;
            if (start && end instanceof Date && start instanceof Date) {
                days = Math.max(1, Math.round((end - start) / (1000 * 60 * 60 * 24)) + 1);
            }
            if (!techGradeDays[tech]) techGradeDays[tech] = {A:0, B:0, C:0, D:0};
            techGradeDays[tech][grade] += days;
        });

        const barLabels = allTechs;
        const barData = allGrades.map(grade =>
            barLabels.map(t => (techGradeDays[t] ? techGradeDays[t][grade] : 0))
        );
        const gradeColors = {'A':'#ef4444', 'B':'#f97316', 'C':'#eab308', 'D':'#22c55e'};
        chartInstances['job-duration-chart'] = new Chart(document.getElementById('job-duration-chart'), {
            type: 'bar',
            data: {
                labels: barLabels,
                datasets: allGrades.map((grade, i)=>({
                    label: grade,
                    data: barData[i],
                    backgroundColor: gradeColors[grade]
                }))
            },
            options: {
                responsive: true,
                plugins: { 
                    legend: { display: false },
                    datalabels: { color: 'white', font: { weight: 'bold' }, formatter: (value) => value > 0 ? value : '' }
                },
                scales: { x: { stacked: true }, y: { stacked: true, beginAtZero: true, title: { display: true, text: "Total Job Days" } } }
            }
        });

        const jobDurationLegendEl = document.getElementById('job-duration-legend');
        if (jobDurationLegendEl) {
            const gradeMeanings = {
                'A': 'Difficult',
                'B': 'Medium',
                'C': 'Easy',
                'D': 'No Challenge'
            };
            let legendHTML = '<div class="flex flex-wrap justify-center gap-x-4 gap-y-1">';
            for (const [grade, meaning] of Object.entries(gradeMeanings)) {
                const color = gradeColors[grade] || '#cccccc';
                legendHTML += `<div class="flex items-center"><span class="w-3 h-3 rounded-sm mr-1.5" style="background-color: ${color};"></span><strong>${grade}:</strong><span class="ml-1">${meaning}</span></div>`;
            }
            legendHTML += '</div>';
            jobDurationLegendEl.innerHTML = legendHTML;
        }


        // 2. Technician Skills (Radar Chart for each tech)
        const allPs = ['P1','P2','P3','P4','P5','P6'];
        allTechs.forEach(tech => {
            createTechSkillCard(tech, filteredReportData);
        });

        // 3. Technician Ranking by Average Score
        const techAvgScores = allTechs.map(tech => {
            const scores = filteredReportData
                .filter(r => r.Technician === tech && r.Score && !isNaN(parseFloat(r.Score)))
                .map(r => parseFloat(r.Score));
            const avg = scores.length ? scores.reduce((a, b) => a + b, 0) / scores.length : 0;
            return { tech, avg };
        }).sort((a,b) => b.avg - a.avg);

        const medals = ['🥇', '🥈', '🥉'];
        const rankingLabels = techAvgScores.map((t, i) => i < 3 ? `${medals[i]} ${t.tech}` : t.tech);

        chartInstances['tech-ranking-chart'] = new Chart(document.getElementById('tech-ranking-chart'), {
            type: 'bar',
            data: {
                labels: rankingLabels,
                datasets: [{
                    label: 'Average Score (%)',
                    data: techAvgScores.map(t => (t.avg / 60) * 100),
                    backgroundColor: techAvgScores.map(t => technicianColors[t.tech] || '#6366f1')
                }]
            },
            options: {
                responsive: true,
                indexAxis: 'y',
                plugins: { 
                    legend: { display: false },
                    datalabels: { 
                        anchor: 'end', 
                        align: 'right', 
                        color: 'black', 
                        font: { weight: 'bold'}, 
                        formatter: (value) => value.toFixed(2) + '%'
                    }
                },
                scales: { 
                    x: { 
                        beginAtZero: true,
                        max: 100,
                        ticks: {
                            callback: function(value) {
                                return value + "%"
                            }
                        }
                    } 
                }
            }
        });

        // 4. Country Visits
        countryTechSelect.innerHTML = `<option value="All">All Technicians</option>` + allTechs.map(tech => `<option value="${tech}">${tech}</option>`).join('');
        const updateCountryPie = () => {
            const selectedTech = countryTechSelect.value;
            let dataPool = selectedTech === 'All' 
                ? filteredReportData.filter(r => allTechs.includes(r.Technician)) 
                : filteredReportData.filter(r => r.Technician === selectedTech);

            const countryData = {};
            dataPool.filter(r => r.Country).forEach(r => {
                countryData[r.Country] = (countryData[r.Country] || 0) + 1;
            });
            const countryLabels = Object.keys(countryData);
            const countryCounts = Object.values(countryData);

            if(chartInstances['country-pie-chart']) chartInstances['country-pie-chart'].destroy();
            chartInstances['country-pie-chart'] = new Chart(document.getElementById('country-pie-chart'), {
                type: 'pie',
                data: {
                    labels: countryLabels,
                    datasets: [{ data: countryCounts, backgroundColor: countryLabels.map((c,i) => colorPalette[i % colorPalette.length]) }]
                },
                options: {
                    responsive: true,
                    plugins: {
                        legend: { display: true, position: 'bottom' },
                        datalabels: { color: 'white', font: { weight: 'bold' }, formatter: (value) => value > 0 ? value : '' },
                        tooltip: { callbacks: { label: ctx => `${ctx.label}: ${ctx.raw} visit(s)` } }
                    }
                }
            });
        }
        updateCountryPie();
        countryTechSelect.onchange = updateCountryPie;

        // 5. Job Types per Technician
        const allTypes = Array.from(new Set(filteredReportData.map(r=>r.Type).filter(Boolean)))
            .sort();
        chartInstances['job-type-chart'] = new Chart(document.getElementById('job-type-chart'), {
            type: 'bar',
            data: {
                labels: allTechs,
                datasets: allTypes.map((type, idx) => ({
                    label: type,
                    data: allTechs.map(tech =>
                        filteredReportData.filter(r => r.Technician === tech && r.Type === type).length
                    ),
                    backgroundColor: colorPalette[idx % colorPalette.length]
                }))
            },
            options: {
                responsive: true,
                plugins: { 
                    legend: { display: true, position: 'bottom' },
                    datalabels: { color: 'white', font: { weight: 'bold' }, formatter: (value) => value > 0 ? value : '' }
                },
                scales: { x: { stacked: true }, y: { stacked: true, beginAtZero: true } }
            }
        });

        // 6. All Technicians Spider Chart
        const allTechJobs = filteredReportData.filter(r => allTechs.includes(r.Technician) && r.Score && !isNaN(parseFloat(r.Score)) && parseFloat(r.Score) > 0);
        const allTechAvgs = allPs.map(p => {
            const pValues = allTechJobs.map(r => parseFloat(r[p])).filter(v => !isNaN(v));
            return pValues.length ? pValues.reduce((a, b) => a + b, 0) / pValues.length : 0;
        });
        chartInstances['all-tech-spider-chart'] = new Chart(document.getElementById('all-tech-spider-chart'), {
            type: 'radar',
            data: {
                labels: allPs,
                datasets: [{
                    label: 'Team Average',
                    data: allTechAvgs.map(a => parseFloat(a.toFixed(2))),
                    backgroundColor: 'rgba(168, 85, 247, 0.15)',
                    borderColor: '#a855f7',
                    pointBackgroundColor: '#a855f7'
                }]
            },
            options: {
                responsive: true,
                plugins: {
                    legend: { display: false },
                    tooltip: { callbacks: { label: (ctx) => `Avg ${ctx.label}: ${ctx.raw.toFixed(2)}/10` } }
                },
                scales: { r: { min: 0, max: 10, beginAtZero: true, ticks: { display: true } } }
            }
        });
    }

    function createTechSkillCard(tech, data) {
        const card = document.createElement('div');
        card.className = 'report-card p-4 rounded-lg shadow-md flex flex-col';
        
        const cardHeader = document.createElement('div');
        cardHeader.className = 'flex items-center mb-2';
        const photoSrc = technicianPhotos[tech];
        if (photoSrc) {
            const img = document.createElement('img');
            img.src = photoSrc;
            img.className = 'w-10 h-10 rounded-full mr-3 object-cover';
            cardHeader.appendChild(img);
        }
        const name = document.createElement('h4');
        name.className = 'font-bold text-md';
        name.textContent = tech;
        cardHeader.appendChild(name);
        
        const filtersDiv = document.createElement('div');
        filtersDiv.className = 'flex gap-2 my-2 no-print';
        
        const gradeSelect = document.createElement('select');
        gradeSelect.className = 'w-full p-1 border rounded-md text-xs';
        const customerSelect = document.createElement('select');
        customerSelect.className = 'w-full p-1 border rounded-md text-xs';
        
        filtersDiv.append(gradeSelect, customerSelect);
        
        const canvas = document.createElement('canvas');
        const summaryDiv = document.createElement('div');
        summaryDiv.className = 'mt-4 text-xs text-gray-600 border-t pt-2';
        
        card.append(cardHeader, filtersDiv, canvas, summaryDiv);
        technicianSkillsContainer.appendChild(card);

        const updateCard = () => {
            const selectedGrade = gradeSelect.value;
            const selectedCustomer = customerSelect.value;
            
            let jobs = data.filter(r =>
                r.Technician === tech &&
                r.Score && !isNaN(parseFloat(r.Score)) && parseFloat(r.Score) > 0
            );

            if (selectedGrade) {
                jobs = jobs.filter(r => (r.Grade || '').toString().toUpperCase() === selectedGrade);
            }
            if (selectedCustomer) {
                jobs = jobs.filter(r => r.Customer === selectedCustomer);
            }
            
            const availableCustomers = [...new Set(jobs.map(j => j.Customer).filter(Boolean))].sort();
            const currentCustomerVal = customerSelect.value;
            customerSelect.innerHTML = `<option value="">All Customers</option>` + availableCustomers.map(c => `<option value="${c}">${c}</option>`).join('');
            customerSelect.value = availableCustomers.includes(currentCustomerVal) ? currentCustomerVal : "";

            const pScores = {};
            const avgs = ['P1','P2','P3','P4','P5','P6'].map(p => {
                const pValues = jobs.map(r => parseFloat(r[p])).filter(v => !isNaN(v));
                const avg = pValues.length ? pValues.reduce((a, b) => a + b, 0) / pValues.length : 0;
                pScores[p] = avg;
                return avg;
            });
            
            const avgScore = jobs.length ? (jobs.reduce((sum, r) => sum + (parseFloat(r.Score) || 0), 0) / jobs.length) : 0;
            
            const chartId = `skill-radar-${tech}`;
            if(chartInstances[chartId]) chartInstances[chartId].destroy();
            
            chartInstances[chartId] = new Chart(canvas, {
                type: 'radar',
                data: {
                    labels: ['P1','P2','P3','P4','P5','P6'],
                    datasets: [{
                        label: `Avg Score: ${(avgScore / 60 * 100).toFixed(2)}%`,
                        data: avgs.map(a => parseFloat(a.toFixed(2))),
                        backgroundColor: 'rgba(59, 130, 246, 0.15)',
                        borderColor: technicianColors[tech] || '#3b82f6',
                        pointBackgroundColor: technicianColors[tech] || '#3b82f6'
                    }]
                },
                options: {
                    responsive: true,
                    plugins: {
                        legend: { display: true, position: 'bottom' },
                        tooltip: { callbacks: { label: ctx => `${ctx.label}: ${ctx.raw.toFixed(2)}/10` } }
                    },
                    scales: { r: { min: 0, max: 10, beginAtZero: true, ticks: { display: true } } }
                }
            });

            summaryDiv.innerHTML = getPreconfiguredSummary(pScores);
        };

        const techJobs = data.filter(r => r.Technician === tech);
        const availableGrades = [...new Set(techJobs.map(j => (j.Grade || '').toString().toUpperCase()).filter(Boolean))].sort();
        gradeSelect.innerHTML = `<option value="">All Grades</option>` + availableGrades.map(g => `<option value="${g}">${g}</option>`).join('');
        
        gradeSelect.onchange = updateCard;
        customerSelect.onchange = updateCard;

        updateCard(); // Initial render
    }


    const updateViews = () => {
        renderCalendar();
        renderSummaryTable();
        renderGanttChart();
        if(!reportView.classList.contains('hidden')) {
            renderReportCharts();
        }
    }
    
    function handlePrint(sectionId) {
        const printStyle = document.createElement('style');
        printStyle.id = 'print-styles';
        printStyle.innerHTML = `
            @media print {
                body * {
                    visibility: hidden;
                }
                #${sectionId}, #${sectionId} * {
                    visibility: visible;
                }
                #${sectionId} {
                    position: absolute;
                    left: 0;
                    top: 0;
                    width: 100%;
                    margin: 0;
                    padding: 1rem;
                    box-shadow: none !important;
                }
            }
        `;
        document.head.appendChild(printStyle);
        window.print();
        document.getElementById('print-styles').remove();
    }

    const updateTime = () => {
        const now = new Date();
        const dateOptions = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
        if (dateDisplayEl) {
            dateDisplayEl.textContent = now.toLocaleDateString('en-GB', dateOptions);
        }
        if (timeDisplayEl) {
            timeDisplayEl.textContent = now.toLocaleTimeString('en-GB');
        }
    };
    
    const setRandomQuote = () => {
        if (quoteEl) {
            const randomIndex = Math.floor(Math.random() * inspirationalQuotes.length);
            quoteEl.textContent = `"${inspirationalQuotes[randomIndex]}"`;
        }
    };

    // --- EVENT LISTENERS ---
    filterContainer.addEventListener('change', (e) => {
        if (e.target.type === 'checkbox') {
            const { filterType, filterValue } = e.target.dataset;
            if (e.target.checked) {
                activeFilters[filterType].add(filterValue);
            } else {
                activeFilters[filterType].delete(filterValue);
            }
            updateFilterBadges();
            updateViews();
        }
    });
    
    document.addEventListener('click', (e) => {
        // Hide dropdowns if clicking outside
        if (!e.target.closest('.filter-group')) {
            document.querySelectorAll('.filter-dropdown').forEach(d => d.classList.add('hidden'));
        }
    });

    clearFiltersBtn.addEventListener('click', () => {
        Object.values(activeFilters).forEach(filterSet => filterSet.clear());
        filterContainer.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = false);
        updateFilterBadges();
        updateViews();
    });

    prevMonthBtn.addEventListener('click', () => { currentDate.setMonth(currentDate.getMonth() - 1); updateViews(); });
    nextMonthBtn.addEventListener('click', () => { currentDate.setMonth(currentDate.getMonth() + 1); updateViews(); });
    fileUploadInput.addEventListener('change', handleFileUpload);
    closeModalBtn.addEventListener('click', hideModal);
    modal.addEventListener('click', (e) => { if (e.target === modal) hideModal(); });
    closeJobDetailModalBtn.addEventListener('click', hideJobDetailModal);
    jobDetailModal.addEventListener('click', (e) => { if (e.target === jobDetailModal) hideJobDetailModal(); });
    
    printSummaryBtn.addEventListener('click', () => handlePrint('summary-section'));
    printCalendarBtn.addEventListener('click', () => handlePrint('calendar-print-section'));
    printGanttBtn.addEventListener('click', () => handlePrint('gantt-print-section'));
    printReportBtn.addEventListener('click', () => handlePrint('report-print-section'));
    reportStartDateInput.addEventListener('change', renderReportCharts);
    reportEndDateInput.addEventListener('change', renderReportCharts);


    calendarTab.addEventListener('click', () => {
        calendarTab.classList.add('active');
        ganttTab.classList.remove('active');
        reportTab.classList.remove('active');
        calendarView.classList.remove('hidden');
        ganttView.classList.add('hidden');
        reportView.classList.add('hidden');
    });

    ganttTab.addEventListener('click', () => {
        ganttTab.classList.add('active');
        calendarTab.classList.remove('active');
        reportTab.classList.remove('active');
        ganttView.classList.remove('hidden');
        calendarView.classList.add('hidden');
               reportView.classList.add('hidden');
    });

    reportTab.addEventListener('click', () => {
        reportTab.classList.add('active');
        calendarTab.classList.remove('active');
        ganttTab.classList.remove('active');
        reportView.classList.remove('hidden');
        calendarView.classList.add('hidden');
        ganttView.classList.add('hidden');
        renderReportCharts();
    });
    
    // Gantt control listeners    ganttStartDateInput.addEventListener('change', renderGanttChart);
    ganttEndDateInput.addEventListener('change', renderGanttChart);
    ganttSortSelect.addEventListener('change', renderGanttChart);
    ganttViewModeRadios.forEach(radio => {
        radio.addEventListener('change', (e) => {
            ganttMode = e.target.value;
            document.getElementById('gantt-sort-controls').style.display = ganttMode === 'customer' ? 'flex' : 'none';
            renderGanttChart();
        });
    });
    
    summaryMonthInput.addEventListener('change', (e) => {
        const [year, month] = e.target.value.split('-');
        const newDate = new Date(Date.UTC(year, month - 1, 1));
        renderSummaryTable(newDate);
    });

    // Initial Setup
    const now = new Date();
    const firstDayOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);
    const lastDayOfMonth = new Date(now.getFullYear(), now.getMonth() + 1, 0);
    const firstDayOfYear = new Date(now.getFullYear(), 0, 1);

    ganttStartDateInput.value = firstDayOfMonth.toISOString().split('T')[0];
    ganttEndDateInput.value = lastDayOfMonth.toISOString().split('T')[0];
    reportStartDateInput.value = firstDayOfYear.toISOString().split('T')[0];
    reportEndDateInput.value = now.toISOString().split('T')[0];
    summaryMonthInput.value = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
    document.getElementById('gantt-sort-controls').style.display = 'none';
    
    updateTime();
    setInterval(updateTime, 1000);
    setRandomQuote();
    renderCalendar();
});
