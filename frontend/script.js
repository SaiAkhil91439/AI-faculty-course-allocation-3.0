document.addEventListener('DOMContentLoaded', () => {
    // Elements
    const numSectionsInput = document.getElementById('numSections');
    const numSubjectsInput = document.getElementById('numSubjects');
    const sectionsPerFacultyInput = document.getElementById('sectionsPerFaculty');
    const maxSubjectsInput = document.getElementById('maxSubjectsPerFaculty');
    const deanMaxSectionsInput = document.getElementById('deanMaxSections');
    const hodMaxSectionsInput = document.getElementById('hodMaxSections');
    const highExpThresholdInput = document.getElementById('highExpThreshold');
    const lblExpThreshold = document.getElementById('lblExpThreshold');

    if (highExpThresholdInput && lblExpThreshold) {
        highExpThresholdInput.addEventListener('input', () => {
            lblExpThreshold.textContent = highExpThresholdInput.value;
        });
    }

    const calcFacultyRequired = document.getElementById('calcFacultyRequired');
    const calcFormula = document.getElementById('calcFormula');
    const divisibilityBadge = document.getElementById('divisibilityBadge');

    const weightPref = document.getElementById('weightPref');
    const weightExp = document.getElementById('weightExp');
    const weightSkill = document.getElementById('weightSkill');
    const weightQual = document.getElementById('weightQual');
    const weightsTotal = document.getElementById('weightsTotal');

    const valPref = document.getElementById('valPref');
    const valExp = document.getElementById('valExp');
    const valSkill = document.getElementById('valSkill');
    const valQual = document.getElementById('valQual');

    const fileCourse = document.getElementById('fileCourse');
    const filePref = document.getElementById('filePref');
    const fileExp = document.getElementById('fileExp');
    const nameCourse = document.getElementById('nameCourse');
    const namePref = document.getElementById('namePref');
    const nameExp = document.getElementById('nameExp');
    const dropzoneCourse = document.getElementById('dropzoneCourse');
    const dropzonePref = document.getElementById('dropzonePref');
    const dropzoneExp = document.getElementById('dropzoneExp');

    const btnValidate = document.getElementById('btnValidate');
    const btnDownloadSample = document.getElementById('btnDownloadSample');
    const btnLogout = document.getElementById('btnLogout');

    if (btnLogout) {
        btnLogout.addEventListener('click', async () => {
            try { await fetch('/api/logout', { method: 'POST' }); }
            finally { window.location.href = '/login'; }
        });
    }

    const allocationForm = document.getElementById('allocationForm');

    const emptyState = document.getElementById('emptyState');
    const loadingState = document.getElementById('loadingState');
    const errorBox = document.getElementById('errorBox');
    const errorMessage = document.getElementById('errorMessage');
    const errorDetails = document.getElementById('errorDetails');
    const resultsContent = document.getElementById('resultsContent');

    // Metrics Display
    const resSections = document.getElementById('resSections');
    const resSubjects = document.getElementById('resSubjects');
    const resFacReq = document.getElementById('resFacReq');
    const resAvgScore = document.getElementById('resAvgScore');

    // Tab Containers
    const matrixTableContainer = document.getElementById('matrixTableContainer');
    const facultyTableContainer = document.getElementById('facultyTableContainer');
    const aiExplanationContent = document.getElementById('aiExplanationContent');
    const traceTimeline = document.getElementById('traceTimeline');

    // 1. Live Calculation of Required Faculty
    function updateCalculatedMetrics() {
        const sections = parseInt(numSectionsInput.value) || 0;
        const perFaculty = parseInt(sectionsPerFacultyInput.value) || 0;

        if (perFaculty > 0 && sections > 0) {
            const required = sections / perFaculty;
            calcFacultyRequired.textContent = `${required} faculty per subject`;
            calcFormula.textContent = `Formula: ${sections} sections / ${perFaculty} sections per faculty`;

            const remainder = sections % perFaculty;
            if (remainder === 0) {
                divisibilityBadge.className = 'divisibility-badge valid';
                divisibilityBadge.innerHTML = '<i class="fa-solid fa-check-circle"></i> Valid Division';
            } else {
                divisibilityBadge.className = 'divisibility-badge invalid';
                divisibilityBadge.innerHTML = `<i class="fa-solid fa-triangle-exclamation"></i> Remainder: ${remainder}`;
            }
        } else {
            calcFacultyRequired.textContent = 'Invalid configuration';
            calcFormula.textContent = 'Sections per faculty must be > 0';
            divisibilityBadge.className = 'divisibility-badge invalid';
            divisibilityBadge.innerHTML = '<i class="fa-solid fa-circle-xmark"></i> Invalid';
        }
    }

    numSectionsInput.addEventListener('input', updateCalculatedMetrics);
    sectionsPerFacultyInput.addEventListener('input', updateCalculatedMetrics);
    updateCalculatedMetrics();

    // 2. Weight Sliders Sync
    function updateWeightDisplays() {
        valPref.textContent = `${weightPref.value}%`;
        valExp.textContent = `${weightExp.value}%`;
        valSkill.textContent = `${weightSkill.value}%`;
        valQual.textContent = `${weightQual.value}%`;

        const total = parseInt(weightPref.value) + parseInt(weightExp.value) + parseInt(weightSkill.value) + parseInt(weightQual.value);
        weightsTotal.textContent = `Total: ${total}%`;
        weightsTotal.style.color = total === 100 ? 'var(--accent)' : 'var(--warning)';
    }

    [weightPref, weightExp, weightSkill, weightQual].forEach(slider => {
        slider.addEventListener('input', updateWeightDisplays);
    });
    updateWeightDisplays();

    // 3. Dropzone File Pickers & Drag-Drop
    function setupDropzone(dropzone, input, nameLabel, defaultText) {
        if (!dropzone || !input) return;
        dropzone.addEventListener('click', (e) => {
            if (e.target.tagName !== 'A') {
                input.click();
            }
        });

        input.addEventListener('change', () => {
            if (input.files.length > 0) {
                nameLabel.textContent = `Selected: ${input.files[0].name}`;
                nameLabel.style.color = '#34d399';
            } else {
                nameLabel.textContent = defaultText;
                nameLabel.style.color = 'var(--accent)';
            }
        });

        ['dragenter', 'dragover'].forEach(eventName => {
            dropzone.addEventListener(eventName, (e) => {
                e.preventDefault();
                dropzone.classList.add('dragover');
            }, false);
        });

        ['dragleave', 'drop'].forEach(eventName => {
            dropzone.addEventListener(eventName, (e) => {
                e.preventDefault();
                dropzone.classList.remove('dragover');
            }, false);
        });

        dropzone.addEventListener('drop', (e) => {
            const dt = e.dataTransfer;
            const files = dt.files;
            if (files.length > 0) {
                input.files = files;
                nameLabel.textContent = `Selected: ${files[0].name}`;
                nameLabel.style.color = '#34d399';
            }
        });
    }

    setupDropzone(dropzoneCourse, fileCourse, nameCourse, 'Using default course_details.xlsx');
    setupDropzone(dropzonePref, filePref, namePref, 'Using default faculty_sample_combined.xlsx');
    setupDropzone(dropzoneExp, fileExp, nameExp, 'Using default faculty_sample_combined.xlsx');

    // 4. Download Sample Excel
    btnDownloadSample.addEventListener('click', async () => {
        try {
            const response = await fetch('/api/sample-data', { method: 'POST' });
            if (response.ok) {
                const blob = await response.blob();
                const downloadUrl = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = downloadUrl;
                a.download = 'faculty_sample_excel_files.zip';
                document.body.appendChild(a);
                a.click();
                a.remove();
            }
        } catch (err) {
            alert('Failed to download sample data package.');
        }
    });

    // Helper: Show UI States
    function showState(stateName) {
        emptyState.classList.add('hidden');
        loadingState.classList.add('hidden');
        errorBox.classList.add('hidden');
        resultsContent.classList.add('hidden');

        if (stateName === 'loading') loadingState.classList.remove('hidden');
        if (stateName === 'error') errorBox.classList.remove('hidden');
        if (stateName === 'results') resultsContent.classList.remove('hidden');
        if (stateName === 'empty') emptyState.classList.remove('hidden');
    }

    // Helper: Render Error Message
    function showError(title, message, details = []) {
        showState('error');
        document.getElementById('errorTitle').textContent = title;
        errorMessage.textContent = message;
        errorDetails.innerHTML = '';
        if (details && details.length > 0) {
            details.forEach(d => {
                const li = document.createElement('li');
                li.textContent = d;
                errorDetails.appendChild(li);
            });
        }
    }

    // 5. Data Validation Trigger
    btnValidate.addEventListener('click', async () => {
        showState('loading');
        document.getElementById('loadingTitle').textContent = 'Validating Excel Files & Designation Rules...';
        document.getElementById('loadingMessage').textContent = 'Checking Dean (4-6 hrs) & HOD (10-12 hrs) workload constraints and High-Exp (2+1 split) requirements.';

        const formData = new FormData();
        formData.append('num_sections', numSectionsInput.value);
        formData.append('num_subjects', numSubjectsInput.value);
        formData.append('sections_per_faculty', sectionsPerFacultyInput.value);
        formData.append('max_subjects_per_faculty', maxSubjectsInput.value);
        formData.append('dean_max_sections', deanMaxSectionsInput.value);
        formData.append('hod_max_sections', hodMaxSectionsInput.value);
        if (highExpThresholdInput) formData.append('high_exp_threshold', highExpThresholdInput.value);

        if (fileCourse && fileCourse.files.length > 0) formData.append('course_file', fileCourse.files[0]);
        if (filePref && filePref.files.length > 0) formData.append('pref_file', filePref.files[0]);
        if (fileExp && fileExp.files.length > 0) formData.append('exp_file', fileExp.files[0]);

        try {
            const response = await fetch('/api/validate', {
                method: 'POST',
                body: formData
            });

            const data = await response.json();
            if (data.success) {
                alert(`Validation Passed!\n\nDeans Found: ${data.summary.deans_count} (Cap: ${data.summary.dean_max_sections} sections)\nHODs Found: ${data.summary.hods_count} (Cap: ${data.summary.hod_max_sections} sections)\nHigh-Exp Faculty (>= ${data.summary.high_exp_threshold} yrs): ${data.summary.high_exp_count}\nCourses Loaded: ${data.summary.courses_loaded}\nTotal Available Faculty: ${data.summary.total_faculty_count}`);
                showState('empty');
            } else {
                showError(data.error_type || 'Validation Failed', data.message, data.details);
            }
        } catch (err) {
            showError('Network / System Error', err.message);
        }
    });

    // 6. Form Submission - Run Allocation Pipeline
    allocationForm.addEventListener('submit', async (e) => {
        e.preventDefault();

        showState('loading');
        document.getElementById('loadingTitle').textContent = 'Running Optimization Engine...';
        document.getElementById('loadingMessage').textContent = 'Solving CP-SAT Model for L-T-P-C contact hours, designation workload limits, and Experienced Faculty 2+1 Section Split.';

        const formData = new FormData();
        formData.append('num_sections', numSectionsInput.value);
        formData.append('num_subjects', numSubjectsInput.value);
        formData.append('sections_per_faculty', sectionsPerFacultyInput.value);
        formData.append('max_subjects_per_faculty', maxSubjectsInput.value);
        formData.append('dean_max_sections', deanMaxSectionsInput.value);
        formData.append('hod_max_sections', hodMaxSectionsInput.value);
        if (highExpThresholdInput) formData.append('high_exp_threshold', highExpThresholdInput.value);

        const totalW = parseInt(weightPref.value) + parseInt(weightExp.value) + parseInt(weightSkill.value) + parseInt(weightQual.value);
        formData.append('weight_pref', parseFloat(weightPref.value) / (totalW || 100));
        formData.append('weight_exp', parseFloat(weightExp.value) / (totalW || 100));
        formData.append('weight_skill', parseFloat(weightSkill.value) / (totalW || 100));
        formData.append('weight_qual', parseFloat(weightQual.value) / (totalW || 100));

        if (fileCourse && fileCourse.files.length > 0) formData.append('course_file', fileCourse.files[0]);
        if (filePref && filePref.files.length > 0) formData.append('pref_file', filePref.files[0]);
        if (fileExp && fileExp.files.length > 0) formData.append('exp_file', fileExp.files[0]);

        try {
            const response = await fetch('/api/allocate', {
                method: 'POST',
                body: formData
            });

            const data = await response.json();
            if (data.success) {
                renderResults(data);
                showState('results');
            } else {
                showError(data.error_type || 'Allocation Error', data.message, data.details);
            }
        } catch (err) {
            showError('Allocation Execution Error', err.message);
        }
    });

    // Render Full Allocation Result Views
    function renderResults(data) {
        const summary = data.summary;
        const allocData = data.allocation_data;

        resSections.textContent = summary.num_sections;
        resSubjects.textContent = summary.num_subjects;
        resFacReq.textContent = summary.faculty_required_per_subject;
        resAvgScore.textContent = `${summary.average_score} / 100`;

        // Render Visual Analytics & Charts
        renderAnalyticsCharts(allocData);

        // 1. Matrix View Table with Role Badges & Color Coding
        const subjects = allocData.subjects;
        const sectionNames = allocData.section_names;
        const matrix = allocData.final_allocation_matrix;
        const facAllocs = allocData.faculty_allocations;

        // Build role lookup for faculty
        const roleMap = {};
        facAllocs.forEach(a => {
            roleMap[a.Lecture_Faculty_Name] = a.Lecture_Role || 'Faculty';
        });

        function getRoleBadge(facName) {
            if (!facName) return '';
            const role = roleMap[facName] || '';
            if (role === 'Dean') return '<span class="role-badge-dean">Dean</span>';
            if (role === 'HOD') return '<span class="role-badge-hod">HOD</span>';
            if (role.includes('Professor') && !role.includes('Assistant')) return '<span class="role-badge-prof">Prof</span>';
            if (role.includes('Assistant')) return '<span class="role-badge-asst">Asst Prof</span>';
            return '';
        }

        let matrixHtml = '<table><thead><tr><th>Section</th>';
        subjects.forEach(s => {
            const ltpc = allocData.course_details && allocData.course_details[s] ? allocData.course_details[s]['L-T-P-C'] : '';
            const hrs = allocData.course_details && allocData.course_details[s] ? allocData.course_details[s]['Total_Hours'] : 5;
            matrixHtml += `<th>${s}<br/><small style="color:#38bdf8; font-weight:600;">[${ltpc}] (${hrs}h)</small></th>`;
        });
        matrixHtml += '</tr></thead><tbody>';

        sectionNames.forEach(sec => {
            matrixHtml += `<tr><td class="td-bold">${sec}</td>`;
            subjects.forEach(s => {
                const cellObj = matrix[sec] ? matrix[sec][s] : null;
                if (typeof cellObj === 'object' && cellObj !== null) {
                    const lBadge = getRoleBadge(cellObj.Lecture_Faculty);
                    matrixHtml += `<td class="td-highlight">
                        <strong style="color:#f8fafc;"><i class="fa-solid fa-chalkboard-user" style="color:#38bdf8;"></i> L:</strong> ${cellObj.Lecture_Faculty} ${lBadge}<br/>
                        <small style="color: #94a3b8;">T/P 1: ${cellObj.TP_Faculty_1}</small><br/>
                        <small style="color: #94a3b8;">T/P 2: ${cellObj.TP_Faculty_2}</small>
                    </td>`;
                } else {
                    matrixHtml += `<td class="td-highlight">${cellObj || 'N/A'}</td>`;
                }
            });
            matrixHtml += '</tr>';
        });
        matrixHtml += '</tbody></table>';
        matrixTableContainer.innerHTML = matrixHtml;

        // 2. Faculty Detail View Table
        let facHtml = `<table>
            <thead>
                <tr>
                    <th>Subject & L-T-P-C</th>
                    <th>Sections</th>
                    <th>Lecture (L) Faculty</th>
                    <th>Designation & Role</th>
                    <th>Course Teaching Exp</th>
                    <th>Teaching Workload (hrs/wk)</th>
                    <th>T/P Faculty 1</th>
                    <th>T/P Faculty 2</th>
                    <th>Pref Rank</th>
                    <th>Score</th>
                </tr>
            </thead>
            <tbody>`;

        facAllocs.forEach(row => {
            const isDeanOrHod = row.Lecture_Role === 'Dean' || row.Lecture_Role === 'HOD';
            const isHighExp = row.Is_High_Experience_Split;
            const roleBadge = isDeanOrHod ? `<span class="badge badge-info" style="background:#059669; color:white;">${row.Lecture_Role}</span>` : `<span class="badge badge-info">${row.Lecture_Role}</span>`;
            const expBadge = isHighExp ? `<br/><small style="color: #f59e0b; font-weight: bold;"><i class="fa-solid fa-star"></i> High-Exp Split (${row.Num_Sections} Secs)</small>` : '';

            facHtml += `<tr ${isHighExp ? 'style="background: rgba(245, 158, 11, 0.08);"' : (isDeanOrHod ? 'style="background: rgba(16, 185, 129, 0.08);"' : '')}>
                <td>
                    <span class="badge badge-info">${row.Subject}</span><br/>
                    <small style="color:#38bdf8; font-weight:600;">L-T-P-C: ${row['L-T-P-C'] || '3-0-2-4'} (${row.Total_Course_Hours || 5} hrs/sec)</small>
                </td>
                <td class="td-bold" style="color: #38bdf8;">${row.Sections}</td>
                <td class="td-bold">${row.Lecture_Faculty_Name} ${expBadge}</td>
                <td>${row.Lecture_Designation} ${roleBadge}</td>
                <td><strong style="color: #60a5fa;">${row.Experience_Years || 0} yrs</strong></td>
                <td>
                    <strong style="color: #34d399;">${row.Assigned_Weekly_Hours} hrs/wk</strong><br/>
                    <small style="color: #94a3b8;">Target: ${row.Target_Hours_Range}</small>
                </td>
                <td>${row.TP_Faculty_1_Name}</td>
                <td>${row.TP_Faculty_2_Name}</td>
                <td><span class="badge-rank">Rank ${row.Preference_Rank}</span></td>
                <td><span class="badge-score">${row.Score}</span></td>
            </tr>`;
        });
        facHtml += '</tbody></table>';
        facultyTableContainer.innerHTML = facHtml;

        // 3. AI Explanation
        aiExplanationContent.innerHTML = data.ai_explanation;

        // 4. Execution Trace Steps
        let traceHtml = '';
        data.agent_steps.forEach(step => {
            traceHtml += `<div class="trace-step">
                <div class="step-num"><i class="fa-solid fa-check"></i></div>
                <div class="step-info">
                    <span class="step-title">Step ${step.step}: ${step.action}</span>
                    <span class="step-desc">${step.details}</span>
                </div>
            </div>`;
        });
        traceTimeline.innerHTML = traceHtml;
    }

    // Chart.js Visual Analytics Engine
    let chartFacultyLoadInstance = null;
    let chartPreferenceSatInstance = null;
    let chartRoleWorkloadInstance = null;

    function renderAnalyticsCharts(allocData) {
        if (typeof Chart === 'undefined') return;
        const facAllocs = allocData.faculty_allocations || [];

        // 1. Faculty Load Distribution Data
        const facultySectionsMap = {};
        facAllocs.forEach(a => {
            const name = a.Lecture_Faculty_Name || 'Faculty';
            const numSecs = a.Num_Sections || 0;
            facultySectionsMap[name] = (facultySectionsMap[name] || 0) + numSecs;
        });

        const facNames = Object.keys(facultySectionsMap);
        const facSections = Object.values(facultySectionsMap);

        const ctxLoad = document.getElementById('chartFacultyLoad');
        if (ctxLoad) {
            if (chartFacultyLoadInstance) chartFacultyLoadInstance.destroy();
            chartFacultyLoadInstance = new Chart(ctxLoad, {
                type: 'bar',
                data: {
                    labels: facNames,
                    datasets: [{
                        label: 'Assigned Sections',
                        data: facSections,
                        backgroundColor: 'rgba(56, 189, 248, 0.65)',
                        borderColor: '#38bdf8',
                        borderWidth: 1.5,
                        borderRadius: 6
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    indexAxis: 'y',
                    plugins: {
                        legend: { display: false },
                        tooltip: {
                            callbacks: {
                                label: (ctx) => ` ${ctx.raw} Sections assigned`
                            }
                        }
                    },
                    scales: {
                        x: {
                            ticks: { color: '#94a3b8', stepSize: 1 },
                            grid: { color: 'rgba(255, 255, 255, 0.05)' }
                        },
                        y: {
                            ticks: { color: '#e2e8f0', font: { size: 11 } },
                            grid: { display: false }
                        }
                    }
                }
            });
        }

        // 2. Preference Satisfaction Doughnut Data
        const rankCounts = { 'Rank 1': 0, 'Rank 2': 0, 'Rank 3': 0, 'Rank 4+': 0 };
        facAllocs.forEach(a => {
            const r = a.Preference_Rank;
            if (r === 1) rankCounts['Rank 1']++;
            else if (r === 2) rankCounts['Rank 2']++;
            else if (r === 3) rankCounts['Rank 3']++;
            else rankCounts['Rank 4+']++;
        });

        const ctxSat = document.getElementById('chartPreferenceSat');
        if (ctxSat) {
            if (chartPreferenceSatInstance) chartPreferenceSatInstance.destroy();
            chartPreferenceSatInstance = new Chart(ctxSat, {
                type: 'doughnut',
                data: {
                    labels: Object.keys(rankCounts),
                    datasets: [{
                        data: Object.values(rankCounts),
                        backgroundColor: [
                            'rgba(16, 185, 129, 0.8)',
                            'rgba(59, 130, 246, 0.8)',
                            'rgba(245, 158, 11, 0.8)',
                            'rgba(239, 68, 68, 0.8)'
                        ],
                        borderColor: 'rgba(15, 23, 42, 0.9)',
                        borderWidth: 2
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: {
                            position: 'bottom',
                            labels: { color: '#cbd5e1', padding: 14, font: { size: 12 } }
                        }
                    }
                }
            });
        }

        // 3. Designation Workload Breakdown Data
        const roleWorkload = {
            'Dean': { totalHrs: 0, count: 0, target: '4-6 hrs' },
            'HOD': { totalHrs: 0, count: 0, target: '10-12 hrs' },
            'Associate Professor': { totalHrs: 0, count: 0, target: '14-16 hrs' },
            'Assistant Professor': { totalHrs: 0, count: 0, target: '16-18 hrs' },
            'Professor': { totalHrs: 0, count: 0, target: '14-16 hrs' }
        };

        facAllocs.forEach(a => {
            const role = a.Lecture_Role || 'Faculty';
            const hrs = a.Assigned_Weekly_Hours || 0;
            if (!roleWorkload[role]) {
                roleWorkload[role] = { totalHrs: 0, count: 0, target: '14-16 hrs' };
            }
            roleWorkload[role].totalHrs += hrs;
            roleWorkload[role].count += 1;
        });

        const activeRoles = Object.keys(roleWorkload).filter(r => roleWorkload[r].count > 0);
        const avgHoursData = activeRoles.map(r => Math.round((roleWorkload[r].totalHrs / roleWorkload[r].count) * 10) / 10);

        const ctxRole = document.getElementById('chartRoleWorkload');
        if (ctxRole) {
            if (chartRoleWorkloadInstance) chartRoleWorkloadInstance.destroy();
            chartRoleWorkloadInstance = new Chart(ctxRole, {
                type: 'bar',
                data: {
                    labels: activeRoles,
                    datasets: [{
                        label: 'Average Weekly Hours',
                        data: avgHoursData,
                        backgroundColor: [
                            'rgba(16, 185, 129, 0.7)',
                            'rgba(139, 92, 246, 0.7)',
                            'rgba(59, 130, 246, 0.7)',
                            'rgba(6, 182, 212, 0.7)',
                            'rgba(245, 158, 11, 0.7)'
                        ],
                        borderColor: [
                            '#10b981', '#8b5cf6', '#3b82f6', '#06b6d4', '#f59e0b'
                        ],
                        borderWidth: 1.5,
                        borderRadius: 6
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: { display: false },
                        tooltip: {
                            callbacks: {
                                label: (ctx) => {
                                    const role = ctx.label;
                                    const target = roleWorkload[role] ? roleWorkload[role].target : 'N/A';
                                    return ` Avg Hours: ${ctx.raw} hrs/wk (Target: ${target})`;
                                }
                            }
                        }
                    },
                    scales: {
                        x: {
                            ticks: { color: '#e2e8f0', font: { size: 11, weight: '600' } },
                            grid: { display: false }
                        },
                        y: {
                            ticks: { color: '#94a3b8' },
                            grid: { color: 'rgba(255, 255, 255, 0.05)' }
                        }
                    }
                }
            });
        }
    }

    // Live Excel Preview Modal Controller
    const btnPreviewExcel = document.getElementById('btnPreviewExcel');
    const previewModal = document.getElementById('previewModal');
    const modalCloseBtn = document.getElementById('modalCloseBtn');
    const previewTabs = document.querySelectorAll('.preview-tab-btn');
    const previewPanes = document.querySelectorAll('.preview-pane');

    if (btnPreviewExcel && previewModal) {
        btnPreviewExcel.addEventListener('click', async () => {
            previewModal.classList.remove('hidden');
            await loadExcelPreview('fileCourse', 'previewCourseContainer', 'Course Details');
            await loadExcelPreview('filePref', 'previewPrefContainer', 'Faculty Preferences');
            await loadExcelPreview('fileExp', 'previewExpContainer', 'Qualifications & Experience');
        });

        modalCloseBtn.addEventListener('click', () => {
            previewModal.classList.add('hidden');
        });
    }

    previewTabs.forEach(tab => {
        tab.addEventListener('click', () => {
            previewTabs.forEach(t => t.classList.remove('active'));
            previewPanes.forEach(p => p.classList.remove('active'));

            tab.classList.add('active');
            const targetId = tab.getAttribute('data-preview');
            const targetPane = document.getElementById(targetId);
            if (targetPane) targetPane.classList.add('active');
        });
    });

    async function loadExcelPreview(fileInputId, containerId, label) {
        const fileInput = document.getElementById(fileInputId);
        const container = document.getElementById(containerId);
        if (!container) return;

        if (fileInput && fileInput.files.length > 0) {
            const file = fileInput.files[0];
            try {
                const buffer = await file.arrayBuffer();
                if (typeof XLSX !== 'undefined') {
                    const workbook = XLSX.read(buffer, { type: 'array' });
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

                    if (jsonData.length > 0) {
                        let html = `<div style="margin-bottom:10px; color:#34d399; font-size:12px; font-weight:600;"><i class="fa-solid fa-file-excel"></i> File: ${file.name} (${jsonData.length - 1} rows)</div><table><thead><tr>`;
                        const headers = jsonData[0];
                        headers.forEach(h => {
                            html += `<th>${h || ''}</th>`;
                        });
                        html += '</tr></thead><tbody>';

                        for (let i = 1; i < Math.min(jsonData.length, 50); i++) {
                            const row = jsonData[i];
                            html += '<tr>';
                            headers.forEach((_, colIdx) => {
                                const val = row[colIdx] !== undefined ? row[colIdx] : '';
                                html += `<td>${val}</td>`;
                            });
                            html += '</tr>';
                        }
                        html += '</tbody></table>';
                        if (jsonData.length > 50) {
                            html += `<p style="text-align:center; margin-top:8px; font-size:11px; color:#94a3b8;">Showing first 50 rows of ${jsonData.length - 1} total rows.</p>`;
                        }
                        container.innerHTML = html;
                        return;
                    }
                }
            } catch (err) {
                container.innerHTML = `<p class="preview-placeholder" style="color:var(--error);">Failed to parse Excel file: ${err.message}</p>`;
                return;
            }
        }
        container.innerHTML = `<p class="preview-placeholder">No custom ${label} file selected. Using system sample Excel file.</p>`;
    }

    // Tab Switching
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
            document.querySelectorAll('.tab-pane').forEach(p => p.classList.remove('active'));

            btn.classList.add('active');
            const tabId = btn.getAttribute('data-tab');
            document.getElementById(tabId).classList.add('active');
        });
    });

    // Live AI Chatbot Widget Controls
    const chatWidgetContainer = document.getElementById('chatWidgetContainer');
    const chatTriggerBtn = document.getElementById('chatTriggerBtn');
    const chatDrawer = document.getElementById('chatDrawer');
    const chatCloseBtn = document.getElementById('chatCloseBtn');
    const chatBody = document.getElementById('chatBody');
    const chatInput = document.getElementById('chatInput');
    const chatSendBtn = document.getElementById('chatSendBtn');
    const chatSuggestions = document.getElementById('chatSuggestions');

    if (chatTriggerBtn && chatDrawer) {
        chatTriggerBtn.addEventListener('click', () => {
            chatDrawer.classList.toggle('hidden');
            if (!chatDrawer.classList.contains('hidden')) {
                chatInput.focus();
            }
        });

        chatCloseBtn.addEventListener('click', () => {
            chatDrawer.classList.add('hidden');
        });
    }

    function formatMarkdown(text) {
        if (!text) return '';
        let formatted = text
            .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
            .replace(/\*(.*?)\*/g, '<em>$1</em>')
            .replace(/`([^`]+)`/g, '<code>$1</code>')
            .replace(/^### (.*$)/gim, '<h4 style="margin:6px 0; color:#38bdf8;">$1</h4>')
            .replace(/^## (.*$)/gim, '<h3 style="margin:8px 0; color:#38bdf8;">$1</h3>')
            .replace(/^# (.*$)/gim, '<h2 style="margin:10px 0; color:#38bdf8;">$1</h2>')
            .replace(/^\> (.*$)/gim, '<blockquote style="border-left:3px solid #38bdf8; padding-left:8px; margin:6px 0; color:#94a3b8;">$1</blockquote>')
            .replace(/^\- (.*$)/gim, '• $1<br/>');

        // Simple Markdown Table parsing
        if (formatted.includes('|')) {
            const lines = formatted.split('\n');
            let inTable = false;
            let tableHtml = '<table>';
            let newLines = [];
            lines.forEach(line => {
                if (line.trim().startsWith('|')) {
                    if (line.includes(':---') || line.includes('---')) return; // Skip separator
                    inTable = true;
                    const cols = line.split('|').filter((_, idx, arr) => idx > 0 && idx < arr.length - 1);
                    tableHtml += '<tr>' + cols.map(c => `<td>${c.trim()}</td>`).join('') + '</tr>';
                } else {
                    if (inTable) {
                        tableHtml += '</table>';
                        newLines.push(tableHtml);
                        tableHtml = '<table>';
                        inTable = false;
                    }
                    newLines.push(line);
                }
            });
            if (inTable) {
                tableHtml += '</table>';
                newLines.push(tableHtml);
            }
            formatted = newLines.join('<br/>');
        }

        return formatted.replace(/\n/g, '<br/>');
    }

    function appendChatMessage(sender, text) {
        const msgDiv = document.createElement('div');
        msgDiv.className = `chat-message ${sender === 'user' ? 'user-msg' : 'assistant-msg'}`;

        const avatar = document.createElement('div');
        avatar.className = 'msg-avatar';
        avatar.innerHTML = sender === 'user' ? '<i class="fa-solid fa-user"></i>' : '<i class="fa-solid fa-robot"></i>';

        const bubble = document.createElement('div');
        bubble.className = 'msg-bubble';
        bubble.innerHTML = sender === 'user' ? text : formatMarkdown(text);

        msgDiv.appendChild(avatar);
        msgDiv.appendChild(bubble);
        chatBody.appendChild(msgDiv);
        chatBody.scrollTop = chatBody.scrollHeight;
    }

    async function sendUserChatMessage(questionText) {
        const query = (questionText || chatInput.value).trim();
        if (!query) return;

        appendChatMessage('user', query);
        chatInput.value = '';

        // Typing indicator
        const typingDiv = document.createElement('div');
        typingDiv.className = 'chat-message assistant-msg';
        typingDiv.id = 'typingIndicator';
        typingDiv.innerHTML = `<div class="msg-avatar"><i class="fa-solid fa-robot"></i></div><div class="msg-bubble"><em>Analyzing allocation data...</em></div>`;
        chatBody.appendChild(typingDiv);
        chatBody.scrollTop = chatBody.scrollHeight;

        try {
            const res = await fetch('/api/chat', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ question: query })
            });
            const data = await res.json();

            const indicator = document.getElementById('typingIndicator');
            if (indicator) indicator.remove();

            if (data.answer) {
                appendChatMessage('assistant', data.answer);
            } else {
                appendChatMessage('assistant', '⚠️ Unable to process chat request.');
            }

            // Update suggestion chips if provided
            if (data.suggested_questions && data.suggested_questions.length > 0) {
                chatSuggestions.innerHTML = '';
                data.suggested_questions.forEach(sq => {
                    const btn = document.createElement('button');
                    btn.className = 'chip-btn';
                    btn.setAttribute('data-query', sq);
                    btn.textContent = sq;
                    chatSuggestions.appendChild(btn);
                });
            }
        } catch (err) {
            const indicator = document.getElementById('typingIndicator');
            if (indicator) indicator.remove();
            appendChatMessage('assistant', '❌ Network error connecting to chat assistant.');
        }
    }

    if (chatSendBtn) {
        chatSendBtn.addEventListener('click', () => sendUserChatMessage());
    }

    if (chatInput) {
        chatInput.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') sendUserChatMessage();
        });
    }

    if (chatSuggestions) {
        chatSuggestions.addEventListener('click', (e) => {
            if (e.target.classList.contains('chip-btn')) {
                const q = e.target.getAttribute('data-query');
                sendUserChatMessage(q);
            }
        });
    }

});

