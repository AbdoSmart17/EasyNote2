// تطبيق إدارة نقاط التلاميذ - ملف JavaScript الموحد
// تم التطوير بواسطة: بقة عبد الوهاب
class StudentGradesApp {
    constructor() {
        this.studentsData = [];
        this.currentSheet = null;
        this.selectedRows = new Set();
        this.allRows = [];
        this.currentSheetName = '';
        
        // أعمدة البيانات
        this.lastNameColumn = -1;
        this.firstNameColumn = -1;
        this.studentRowsStart = -1;
        this.studentRowsEnd = -1;
        
        // أعمدة العلامات
        this.notebookColumn = -1;
        this.homeworkColumn = -1;
        this.behaviorColumn = -1;
        this.participationColumn = -1;
        this.continuousColumn = -1;
        this.testColumn = -1;
        this.examColumn = -1;
        
        // إعدادات التقديرات
        this.gradeSettings = [
            { min: 18, max: 20, comment: 'نتائج ممتازة ومرضية واصل' },
            { min: 16, max: 17.99, comment: 'نتائج جيدة و مرضية واصل' },
            { min: 14, max: 15.99, comment: 'واصل الاجتهاد و المثابرة' },
            { min: 12, max: 13.99, comment: 'نتائج مقبولة بإمكانك تحسينها' },
            { min: 10, max: 11.99, comment: 'بمقدورك تحقيق نتائج أفضل' },
            { min: 7, max: 9.99, comment: 'ينقصك الحرص و التركيز' },
            { min: 0, max: 6.99, comment: 'احذر التهاون' }
        ];
        
        // حالات التحديد
        this.startRowIndex = -1;
        this.dragSelecting = false;
        
        // خصائص المعالجة الشاملة
        this.allProcessedSheets = new Map(); // تخزين جميع الصفحات المعالجة
        this.currentWorkbook = null; // تخزين ملف Excel الأصلي
        
        this.init();
    }

    init() {
        console.log('🚀 تهيئة تطبيق إدارة نقاط التلاميذ...');
        this.loadGradeSettings();
        this.setupEventListeners();
        this.setupPWA();
        this.finalizeInit();
        this.showToast('تم تحميل التطبيق بنجاح', 'success');
    }

    setupEventListeners() {
        // أحداث التبويبات
        document.querySelectorAll('.nav-tab').forEach(tab => {
            tab.addEventListener('click', (e) => {
                const tabName = e.currentTarget.dataset.tab;
                this.showTab(tabName);
            });
        });

        // أحداث الاستيراد
        document.getElementById('uploadTrigger').addEventListener('click', () => {
            document.getElementById('fileInput').click();
        });
        document.getElementById('fileInput').addEventListener('change', (e) => this.handleFile(e));
        document.getElementById('sheetSelect').addEventListener('change', () => this.loadSheet());

        // أحداث التحديد
        document.getElementById('selectAll').addEventListener('click', () => this.selectAllRows());
        document.getElementById('deselectAll').addEventListener('click', () => this.deselectAllRows());
        document.getElementById('selectStudents').addEventListener('click', () => this.selectStudentRows());
        document.getElementById('invertSelection').addEventListener('click', () => this.invertSelection());
        document.getElementById('confirmSelection').addEventListener('click', () => this.confirmSelection());
        document.getElementById('clearSelection').addEventListener('click', () => this.clearSelection());

        // أحداث النقاط
        document.getElementById('gradeType').addEventListener('change', () => this.updateGradeInputs());
        document.getElementById('applyToAll').addEventListener('click', () => this.applyGradeToAll());
        document.getElementById('searchBox').addEventListener('input', () => this.searchStudents());
        document.getElementById('sortByName').addEventListener('click', () => this.sortStudents('name'));
        document.getElementById('sortByAverage').addEventListener('click', () => this.sortStudents('average'));
        document.getElementById('resetGrades').addEventListener('click', () => this.resetGrades());
        document.getElementById('exportExcel').addEventListener('click', () => this.exportToExcel());

        // أحداث الإحصائيات
        document.getElementById('exportStats').addEventListener('click', () => this.exportStatistics());

        // أحداث الإعدادات
        document.getElementById('saveSettings').addEventListener('click', () => this.saveGradeSettings());
        document.getElementById('resetSettings').addEventListener('click', () => this.resetGradeSettings());
        document.getElementById('exportSettings').addEventListener('click', () => this.exportSettings());
        document.getElementById('importSettingsInput').addEventListener('change', (e) => this.importSettings(e));
        document.getElementById('resetAllData').addEventListener('click', () => this.resetAllData());
        document.getElementById('resetGradeSettingsBtn').addEventListener('click', () => this.resetGradeSettings());

        // أحداث المعالجة الشاملة
        document.getElementById('processAllSheets').addEventListener('click', () => this.processAllSheets());
        document.getElementById('exportAllSheets').addEventListener('click', () => this.exportAllProcessedSheets());
        document.getElementById('previewAllSheets').addEventListener('click', () => this.previewAllSheets());

        // منع إغلاق الصفحة مع بيانات غير محفوظة
        window.addEventListener('beforeunload', (e) => {
            if (this.studentsData.length > 0) {
                e.preventDefault();
                e.returnValue = 'لديك بيانات غير محفوظة. هل تريد المغادرة حقاً؟';
            }
        });
    }

    // إدارة التبويبات
    showTab(tabName) {
        document.querySelectorAll('.tab-content').forEach(tab => tab.classList.remove('active'));
        document.querySelectorAll('.nav-tab').forEach(tab => tab.classList.remove('active'));

        document.getElementById(tabName + '-tab').classList.add('active');
        document.querySelector(`[data-tab="${tabName}"]`).classList.add('active');

        if (tabName === 'statistics') {
            this.updateStatistics();
        }
        
        this.showToast(`تم التبديل إلى ${this.getTabName(tabName)}`, 'success');
    }

    getTabName(tabId) {
        const tabNames = {
            'import': 'استيراد البيانات',
            'grades': 'إدخال النقاط',
            'statistics': 'الإحصائيات',
            'settings': 'الإعدادات'
        };
        return tabNames[tabId] || tabId;
    }

    // نظام الإشعارات
    showToast(message, type = 'success') {
        const toastContainer = document.getElementById('toastContainer');
        const toast = document.createElement('div');
        toast.className = `toast ${type}`;
        
        const icons = {
            success: '✅',
            error: '❌',
            warning: '⚠️',
            info: 'ℹ️'
        };
        
        toast.innerHTML = `
            <span class="toast-icon">${icons[type] || '💡'}</span>
            <span class="toast-message">${message}</span>
        `;
        
        toastContainer.appendChild(toast);
        
        setTimeout(() => {
            toast.style.animation = 'slideDown 0.3s ease reverse forwards';
            setTimeout(() => toast.remove(), 300);
        }, 3000);
    }

    showLoading() {
        document.getElementById('loadingSpinner').style.display = 'block';
    }

    hideLoading() {
        document.getElementById('loadingSpinner').style.display = 'none';
    }

    handleError(error, context) {
        console.error(`❌ خطأ في ${context}:`, error);
        this.showToast(`حدث خطأ: ${error.message}`, 'error');
    }

    // أدوات مساعدة
    getColumnName(columnIndex) {
        let result = '';
        let index = columnIndex;
        
        do {
            result = String.fromCharCode(65 + (index % 26)) + result;
            index = Math.floor(index / 26) - 1;
        } while (index >= 0);
        
        return result;
    }

    getGradeComment(average) {
        for (const setting of this.gradeSettings) {
            if (average >= setting.min && average <= setting.max) {
                return setting.comment;
            }
        }
        return 'غير محدد';
    }

    calculateAverage(index) {
        const student = this.studentsData[index];
        
        student.continuous = (student.notebook + student.homework + student.behavior + student.participation) / 4;
        student.average = (student.continuous + student.test + (student.exam * 3)) / 5;
        student.comment = this.getGradeComment(student.average);
    }

    // PWA - تطبيق الويب التقدمي
    setupPWA() {
        if ('serviceWorker' in navigator) {
            window.addEventListener('load', () => {
                navigator.serviceWorker.register('/sw.js')
                    .then((registration) => {
                        console.log('✅ ServiceWorker مسجل بنجاح: ', registration.scope);
                    })
                    .catch((error) => {
                        console.log('❌ فشل تسجيل ServiceWorker: ', error);
                    });
            });
        }
        
        this.addInstallButton();
    }

    addInstallButton() {
        // يمكن إضافة زر التثبيت هنا إذا لزم الأمر
    }

    // استيراد البيانات من Excel
    handleFile(event) {
        const file = event.target.files[0];
        if (!file) return;

        this.showLoading();
        
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = e.target.result;
                const workbook = XLSX.read(data, {type: 'binary'});
                
                // حفظ الملف للاستخدام لاحقاً
                this.currentWorkbook = workbook;
                
                // تعبئة محدد الصفحات الفردي
                const sheetSelect = document.getElementById('sheetSelect');
                sheetSelect.innerHTML = '<option value="">-- اختر القسم --</option>';
                workbook.SheetNames.forEach(name => {
                    const option = document.createElement('option');
                    option.value = name;
                    option.textContent = name;
                    sheetSelect.appendChild(option);
                });

                document.getElementById('sheetSelector').style.display = 'block';
                
                // إظهار قسم المعالجة الشاملة
                document.getElementById('bulkProcessingSection').style.display = 'block';
                
                this.hideLoading();
                this.showToast('تم تحميل الملف بنجاح - يمكنك معالجة صفحات فردية أو جميع الصفحات', 'success');
            } catch (error) {
                this.hideLoading();
                this.handleError(error, 'تحميل الملف');
            }
        };
        
        reader.onerror = () => {
            this.hideLoading();
            this.showToast('خطأ في قراءة الملف', 'error');
        };
        
        reader.readAsBinaryString(file);
    }

    loadSheet() {
        const sheetName = document.getElementById('sheetSelect').value;
        if (!sheetName || !this.currentWorkbook) return;

        this.showLoading();
        
        try {
            this.currentSheetName = sheetName;
            const worksheet = this.currentWorkbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, {header: 1, defval: ""});
            
            this.currentSheet = jsonData;
            this.allRows = jsonData;
            
            this.findNameColumnsAndStudentRowsAndGrades(jsonData);
            this.displayPreview(jsonData);
            document.getElementById('previewSection').style.display = 'block';
            
            this.hideLoading();
            this.showToast(`تم تحميل الصفحة: ${sheetName}`, 'success');
        } catch (error) {
            this.hideLoading();
            this.handleError(error, 'تحميل الصفحة');
        }
    }

    findNameColumnsAndStudentRowsAndGrades(data) {
        // إعادة تعيين جميع الأعمدة
        const columns = [
            'lastNameColumn', 'firstNameColumn', 'notebookColumn', 'homeworkColumn',
            'behaviorColumn', 'participationColumn', 'continuousColumn', 'testColumn', 'examColumn'
        ];
        columns.forEach(col => this[col] = -1);
        
        this.studentRowsStart = -1;
        this.studentRowsEnd = -1;

        const keywords = {
            lastNameColumn: ['اللقب', 'لقب', 'اسم العائلة', 'Nom', 'Last Name', 'Surname'],
            firstNameColumn: ['الاسم', 'اسم', 'الاسم الشخصي', 'Prénom', 'First Name', 'Name'],
            notebookColumn: ['الكراس', 'كراس', 'Cahier', 'Notebook'],
            homeworkColumn: ['الواجبات', 'واجبات', 'Devoirs', 'Homework'],
            behaviorColumn: ['السلوك', 'سلوك', 'Comportement', 'Behavior'],
            participationColumn: ['المشاركة', 'مشاركة', 'Participation'],
            continuousColumn: ['التقويم المستمر', 'تقويم مستمر','معدل تقويم النشاطات','التقويم', 'تقويم', 'Continu', 'Continuous'],
            testColumn: ['الفرض', 'فرض', 'Devoir', 'Test'],
            examColumn: ['الاختبار', 'اختبار','الإختبار','إختبار','الامتحان' , 'امتحان','Examen', 'Exam']
        };
        
        let headerRowIndex = -1;
        
        for (let rowIndex = 0; rowIndex < Math.min(20, data.length); rowIndex++) {
            const row = data[rowIndex];
            
            for (let colIndex = 0; colIndex < row.length; colIndex++) {
                const cellValue = String(row[colIndex] || '').trim();
                
                for (const [column, keywordList] of Object.entries(keywords)) {
                    if (this[column] === -1 && keywordList.some(keyword => cellValue.includes(keyword))) {
                        this[column] = colIndex;
                        headerRowIndex = rowIndex;
                    }
                }
            }
            
            if (headerRowIndex !== -1) break;
        }
        
        if (headerRowIndex !== -1) {
            this.studentRowsStart = headerRowIndex + 1;
            this.studentRowsEnd = data.length - 1;
            
            for (let i = this.studentRowsStart; i < data.length; i++) {
                const row = data[i];
                const hasData = row && (
                    (this.lastNameColumn !== -1 && row[this.lastNameColumn] && String(row[this.lastNameColumn]).trim()) ||
                    (this.firstNameColumn !== -1 && row[this.firstNameColumn] && String(row[this.firstNameColumn]).trim())
                );
                
                if (!hasData) {
                    this.studentRowsEnd = i - 1;
                    break;
                }
            }
        }
        
        this.displayColumnAndStudentAndGradesInfo();
    }

    displayColumnAndStudentAndGradesInfo() {
        const columnInfo = document.getElementById('columnInfo');
        const studentRowsInfo = document.getElementById('studentRowsInfo');
        
        let infoHTML = '<strong>تم التعرف على الأعمدة التالية:</strong><br>';
        
        if (this.lastNameColumn !== -1) {
            infoHTML += `- عمود اللقب: العمود ${this.lastNameColumn + 1} (${this.getColumnName(this.lastNameColumn)})<br>`;
        } else {
            infoHTML += '- لم يتم العثور على عمود اللقب<br>';
        }
        
        if (this.firstNameColumn !== -1) {
            infoHTML += `- عمود الاسم: العمود ${this.firstNameColumn + 1} (${this.getColumnName(this.firstNameColumn)})<br>`;
        } else {
            infoHTML += '- لم يتم العثور على عمود الاسم<br>';
        }
        
        // عرض معلومات العلامات
        const gradeColumns = [
            { col: this.notebookColumn, name: 'الكراس' },
            { col: this.homeworkColumn, name: 'الواجبات' },
            { col: this.behaviorColumn, name: 'السلوك' },
            { col: this.participationColumn, name: 'المشاركة' },
            { col: this.continuousColumn, name: 'التقويم المستمر' },
            { col: this.testColumn, name: 'الفرض' },
            { col: this.examColumn, name: 'الاختبار' }
        ];
        
        const partialGrades = gradeColumns.slice(0, 4).filter(col => col.col !== -1);
        const totalGrades = gradeColumns.slice(4).filter(col => col.col !== -1);
        
        if (partialGrades.length > 0) {
            infoHTML += '<strong>العلامات الجزئية:</strong><br>';
            partialGrades.forEach(grade => {
                infoHTML += `- ${grade.name}: العمود ${grade.col + 1}<br>`;
            });
        }
        
        if (totalGrades.length > 0) {
            infoHTML += '<strong>العلامات الإجمالية:</strong><br>';
            totalGrades.forEach(grade => {
                infoHTML += `- ${grade.name}: العمود ${grade.col + 1}<br>`;
            });
        }
        
        columnInfo.innerHTML = infoHTML;
        columnInfo.style.display = 'block';
        
        if (this.studentRowsStart !== -1 && this.studentRowsEnd !== -1) {
            const studentCount = this.studentRowsEnd - this.studentRowsStart + 1;
            studentRowsInfo.innerHTML = `
                <strong>تم التعرف على صفوف التلاميذ:</strong><br>
                - بداية التلاميذ: الصف ${this.studentRowsStart + 1}<br>
                - نهاية التلاميذ: الصف ${this.studentRowsEnd + 1}<br>
                - عدد التلاميذ: ${studentCount} تلميذ
            `;
            studentRowsInfo.style.display = 'block';
        } else {
            studentRowsInfo.innerHTML = '<strong>لم يتم التعرف على صفوف التلاميذ تلقائياً.</strong>';
            studentRowsInfo.style.display = 'block';
        }
    }

    displayPreview(data) {
        const table = document.getElementById('previewTable');
        table.innerHTML = '';

        const thead = document.createElement('thead');
        const headerRow = document.createElement('tr');
        
        let headerHTML = '<th>الصف</th>';
        const columnCount = Math.min(8, data[0] ? data[0].length : 8);
        
        for (let i = 0; i < columnCount; i++) {
            let columnName = `العمود ${i + 1}`;
            if (i === this.lastNameColumn) columnName += ' (اللقب)';
            if (i === this.firstNameColumn) columnName += ' (الاسم)';
            if (i === this.continuousColumn) columnName += ' (التقويم)';
            if (i === this.testColumn) columnName += ' (الفرض)';
            if (i === this.examColumn) columnName += ' (الاختبار)';
            headerHTML += `<th>${columnName}</th>`;
        }
        
        headerRow.innerHTML = headerHTML;
        thead.appendChild(headerRow);
        table.appendChild(thead);

        const tbody = document.createElement('tbody');
        
        data.forEach((row, index) => {
            const tr = document.createElement('tr');
            tr.classList.add('selectable-row');
            tr.dataset.rowIndex = index;
            
            const isStudentRow = (index >= this.studentRowsStart && index <= this.studentRowsEnd);
            const hasStudentData = isStudentRow && 
                                  (this.lastNameColumn !== -1 && this.firstNameColumn !== -1 && 
                                   row[this.lastNameColumn] && row[this.firstNameColumn] && 
                                   String(row[this.lastNameColumn]).trim() && String(row[this.firstNameColumn]).trim());
            
            if (hasStudentData) {
                tr.style.background = '#f0f8ff';
            } else if (isStudentRow) {
                tr.style.background = '#fff3cd';
            }
            
            let rowHTML = `<td style="font-weight: bold; color: #007bff;">${index + 1}</td>`;
            for (let i = 0; i < columnCount; i++) {
                let cellValue = row[i] || '';
                let cellStyle = '';
                
                if (i === this.lastNameColumn || i === this.firstNameColumn) cellStyle = 'background: #e8f5e8;';
                if (i === this.continuousColumn || i === this.testColumn || i === this.examColumn) cellStyle = 'background: #e3f2fd;';
                
                rowHTML += `<td style="${cellStyle}">${cellValue}</td>`;
            }
            
            tr.innerHTML = rowHTML;
            
            tr.addEventListener('mousedown', (e) => this.startSelection(e));
            tr.addEventListener('mouseenter', (e) => this.continueSelection(e));
            tr.addEventListener('mouseup', (e) => this.endSelection(e));
            tbody.appendChild(tr);
        });
        
        table.appendChild(tbody);
        document.getElementById('totalRowsCount').textContent = `إجمالي الصفوف: ${data.length}`;
        this.selectedRows.clear();
        this.updateSelectionCount();
    }

    // وظائف التحديد
    startSelection(event) {
        event.preventDefault();
        const row = event.currentTarget;
        const rowIndex = parseInt(row.dataset.rowIndex);
        
        if (event.ctrlKey) {
            if (this.selectedRows.has(rowIndex)) {
                this.selectedRows.delete(rowIndex);
                row.classList.remove('selected');
            } else {
                this.selectedRows.add(rowIndex);
                row.classList.add('selected');
            }
        } else {
            if (!event.shiftKey) {
                this.clearSelection();
            }
            
            this.startRowIndex = rowIndex;
            this.selectedRows.add(rowIndex);
            row.classList.add('selected');
            this.dragSelecting = true;
        }
        this.updateSelectionCount();
    }

    continueSelection(event) {
        if (!this.dragSelecting || this.startRowIndex === -1) return;
        
        const row = event.currentTarget;
        const rowIndex = parseInt(row.dataset.rowIndex);
        
        document.querySelectorAll('.selectable-row').forEach(r => {
            const rIndex = parseInt(r.dataset.rowIndex);
            if (this.selectedRows.has(rIndex) && rIndex !== this.startRowIndex) {
                this.selectedRows.delete(rIndex);
                r.classList.remove('selected');
            }
        });
        
        const minIndex = Math.min(this.startRowIndex, rowIndex);
        const maxIndex = Math.max(this.startRowIndex, rowIndex);
        
        this.selectedRows.clear();
        this.selectedRows.add(this.startRowIndex);
        
        for (let i = minIndex; i <= maxIndex; i++) {
            this.selectedRows.add(i);
            const rowElement = document.querySelector(`[data-row-index="${i}"]`);
            if (rowElement) {
                rowElement.classList.add('selected');
            }
        }
        this.updateSelectionCount();
    }

    endSelection(event) {
        this.dragSelecting = false;
        this.updateSelectionCount();
    }

    clearSelection() {
        document.querySelectorAll('.selectable-row').forEach(r => r.classList.remove('selected'));
        this.selectedRows.clear();
        this.startRowIndex = -1;
        this.dragSelecting = false;
        this.updateSelectionCount();
    }

    selectAllRows() {
        document.querySelectorAll('.selectable-row').forEach(row => {
            const rowIndex = parseInt(row.dataset.rowIndex);
            this.selectedRows.add(rowIndex);
            row.classList.add('selected');
        });
        this.updateSelectionCount();
    }

    deselectAllRows() {
        this.clearSelection();
    }

    invertSelection() {
        document.querySelectorAll('.selectable-row').forEach(row => {
            const rowIndex = parseInt(row.dataset.rowIndex);
            if (this.selectedRows.has(rowIndex)) {
                this.selectedRows.delete(rowIndex);
                row.classList.remove('selected');
            } else {
                this.selectedRows.add(rowIndex);
                row.classList.add('selected');
            }
        });
        this.updateSelectionCount();
    }

    selectStudentRows() {
        this.clearSelection();
        
        if (this.studentRowsStart === -1 || this.studentRowsEnd === -1) {
            this.showToast('لم يتم التعرف على صفوف التلاميذ. يرجى التحديد يدوياً.', 'warning');
            return;
        }
        
        document.querySelectorAll('.selectable-row').forEach(row => {
            const rowIndex = parseInt(row.dataset.rowIndex);
            
            if (rowIndex >= this.studentRowsStart && rowIndex <= this.studentRowsEnd) {
                this.selectedRows.add(rowIndex);
                row.classList.add('selected');
            }
        });
        
        this.updateSelectionCount();
        
        if (this.selectedRows.size > 0) {
            this.showToast(`تم تحديد ${this.selectedRows.size} صف تلقائياً يحتوي على بيانات التلاميذ`, 'success');
        } else {
            this.showToast('لم يتم العثور على صفوف تحتوي على بيانات التلاميذ. يرجى التحديد يدوياً.', 'warning');
        }
    }

    updateSelectionCount() {
        const count = this.selectedRows.size;
        const countElement = document.getElementById('selectionCount');
        if (count === 0) {
            countElement.textContent = 'لم يتم التحديد';
            countElement.style.color = '#6c757d';
        } else {
            countElement.textContent = `تم تحديد ${count} صف`;
            countElement.style.color = '#007bff';
        }
    }

    confirmSelection() {
        if (this.selectedRows.size === 0) {
            this.showToast('يرجى تحديد الصفوف التي تحتوي على أسماء التلاميذ', 'warning');
            return;
        }

        this.studentsData = [];
        const validStudents = [];
        
        this.selectedRows.forEach(rowIndex => {
            const rowData = this.currentSheet[rowIndex];
            if (rowData && 
                (this.lastNameColumn === -1 || rowData[this.lastNameColumn]) && 
                (this.firstNameColumn === -1 || rowData[this.firstNameColumn])) {
                
                const lastName = this.lastNameColumn !== -1 ? String(rowData[this.lastNameColumn]).trim() : '';
                const firstName = this.firstNameColumn !== -1 ? String(rowData[this.firstNameColumn]).trim() : '';
                
                if (lastName && firstName) {
                    let notebook = parseFloat(rowData[this.notebookColumn]) || 0;
                    let homework = parseFloat(rowData[this.homeworkColumn]) || 0;
                    let behavior = parseFloat(rowData[this.behaviorColumn]) || 0;
                    let participation = parseFloat(rowData[this.participationColumn]) || 0;
                    let continuous = parseFloat(rowData[this.continuousColumn]) || 0;
                    let test = parseFloat(rowData[this.testColumn]) || 0;
                    let exam = parseFloat(rowData[this.examColumn]) || 0;
                    
                    if (continuous > 0 && notebook === 0 && homework === 0 && behavior === 0 && participation === 0) {
                        notebook = homework = behavior = participation = continuous;
                    }
                    
                    validStudents.push({
                        name: `${lastName} ${firstName}`,
                        lastName: lastName,
                        firstName: firstName,
                        notebook: notebook,
                        homework: homework,
                        behavior: behavior,
                        participation: participation,
                        test: test,
                        exam: exam,
                        continuous: 0,
                        average: 0,
                        comment: ''
                    });
                }
            }
        });

        if (validStudents.length === 0) {
            this.showToast('لم يتم العثور على بيانات صالحة للتلاميذ.', 'error');
            return;
        }

        this.studentsData = validStudents;
        
        this.studentsData.forEach((student, index) => {
            this.calculateAverage(index);
        });
        
        this.updateGradesTable();
        
        const message = `تم تحديد ${this.studentsData.length} تلميذ بنجاح!\n\nمن أصل ${this.selectedRows.size} صف محدد، تم قبول ${this.studentsData.length} تلميذ يحتوون على بيانات صالحة.`;
        this.showToast(message, 'success');
        this.showTab('grades');
    }

    // إدارة النقاط
    updateGradesTable() {
        const tbody = document.getElementById('gradesTableBody');
        if (!tbody) return;
        
        tbody.innerHTML = '';

        this.studentsData.forEach((student, index) => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${student.name}</td>
                <td><input type="number" class="grade-input" value="${student.notebook}" min="0" max="20" step="0.5" onchange="app.updateStudentGrade(${index}, 'notebook', this.value)"></td>
                <td><input type="number" class="grade-input" value="${student.homework}" min="0" max="20" step="0.5" onchange="app.updateStudentGrade(${index}, 'homework', this.value)"></td>
                <td><input type="number" class="grade-input" value="${student.behavior}" min="0" max="20" step="0.5" onchange="app.updateStudentGrade(${index}, 'behavior', this.value)"></td>
                <td><input type="number" class="grade-input" value="${student.participation}" min="0" max="20" step="0.5" onchange="app.updateStudentGrade(${index}, 'participation', this.value)"></td>
                <td class="average-display">${student.continuous.toFixed(2)}</td>
                <td><input type="number" class="grade-input" value="${student.test}" min="0" max="20" step="0.5" onchange="app.updateStudentGrade(${index}, 'test', this.value)"></td>
                <td><input type="number" class="grade-input" value="${student.exam}" min="0" max="20" step="0.5" onchange="app.updateStudentGrade(${index}, 'exam', this.value)"></td>
                <td class="average-display">${student.average.toFixed(2)}</td>
                <td class="grade-comment">${student.comment}</td>
            `;
            tbody.appendChild(tr);
        });
    }

    updateStudentGrade(index, field, value) {
        const numValue = parseFloat(value) || 0;
        this.studentsData[index][field] = Math.max(0, Math.min(20, numValue));
        this.calculateAverage(index);
        this.updateGradesTable();
        this.showToast(`تم تحديث نقاط ${this.studentsData[index].name}`, 'success');
    }

    updateGradeInputs() {
        const gradeType = document.getElementById('gradeType').value;
        const continuousDiv = document.getElementById('continuousGrades');
        const singleDiv = document.getElementById('singleGrade');
        const label = document.getElementById('singleGradeLabel');

        if (gradeType === 'continuous') {
            continuousDiv.style.display = 'grid';
            singleDiv.style.display = 'none';
        } else {
            continuousDiv.style.display = 'none';
            singleDiv.style.display = 'block';
            label.textContent = gradeType === 'test' ? 'الفرض' : 'الإختبار';
        }
    }

    applyGradeToAll() {
        if (this.studentsData.length === 0) {
            this.showToast('لا توجد بيانات تلاميذ', 'warning');
            return;
        }

        const gradeType = document.getElementById('gradeType').value;
        
        if (gradeType === 'continuous') {
            const notebook = parseFloat(document.getElementById('notebook').value) || 0;
            const homework = parseFloat(document.getElementById('homework').value) || 0;
            const behavior = parseFloat(document.getElementById('behavior').value) || 0;
            const participation = parseFloat(document.getElementById('participation').value) || 0;
            
            this.studentsData.forEach((student, index) => {
                student.notebook = notebook;
                student.homework = homework;
                student.behavior = behavior;
                student.participation = participation;
                this.calculateAverage(index);
            });
            
            this.showToast(`تم تطبيق العلامات المستمرة على جميع التلاميذ`, 'success');
        } else {
            const gradeValue = parseFloat(document.getElementById('singleGradeInput').value) || 0;
            const field = gradeType === 'test' ? 'test' : 'exam';
            
            this.studentsData.forEach((student, index) => {
                student[field] = gradeValue;
                this.calculateAverage(index);
            });
            
            const gradeName = gradeType === 'test' ? 'الفرض' : 'الاختبار';
            this.showToast(`تم تطبيق ${gradeName} على جميع التلاميذ`, 'success');
        }
        
        this.updateGradesTable();
    }

    searchStudents() {
        const searchTerm = document.getElementById('searchBox').value.toLowerCase();
        const rows = document.querySelectorAll('#gradesTableBody tr');
        
        let visibleCount = 0;
        rows.forEach(row => {
            const name = row.cells[0].textContent.toLowerCase();
            const isVisible = name.includes(searchTerm);
            row.style.display = isVisible ? '' : 'none';
            if (isVisible) visibleCount++;
        });
        
        const searchInfo = document.getElementById('searchInfo');
        if (searchTerm) {
            searchInfo.textContent = `عرض ${visibleCount} من أصل ${rows.length} تلميذ`;
            searchInfo.style.display = 'block';
        } else {
            searchInfo.style.display = 'none';
        }
    }

    sortStudents(criteria) {
        if (this.studentsData.length === 0) {
            this.showToast('لا توجد بيانات للترتيب', 'warning');
            return;
        }

        if (criteria === 'name') {
            this.studentsData.sort((a, b) => a.name.localeCompare(b.name, 'ar'));
            this.showToast('تم الترتيب حسب الأسماء', 'success');
        } else if (criteria === 'average') {
            this.studentsData.sort((a, b) => b.average - a.average);
            this.showToast('تم الترتيب حسب المعدلات', 'success');
        }
        this.updateGradesTable();
    }

    resetGrades() {
        if (this.studentsData.length === 0) {
            this.showToast('لا توجد بيانات لإعادة التعيين', 'warning');
            return;
        }

        if (confirm('هل أنت متأكد من إعادة تعيين جميع النقاط؟ سيتم فقدان جميع البيانات الحالية.')) {
            this.studentsData.forEach((student, index) => {
                student.notebook = 0;
                student.homework = 0;
                student.behavior = 0;
                student.participation = 0;
                student.test = 0;
                student.exam = 0;
                this.calculateAverage(index);
            });
            this.updateGradesTable();
            this.showToast('تم إعادة تعيين جميع النقاط', 'success');
        }
    }

    exportToExcel() {
        if (this.studentsData.length === 0) {
            this.showToast('لا توجد بيانات للتصدير', 'warning');
            return;
        }

        this.showLoading();
        
        try {
            const exportData = [
                ['اللقب والاسم', 'الكراس', 'الواجبات', 'السلوك', 'المشاركة', 'التقويم المستمر', 'الفرض', 'الاختبار', 'المعدل', 'التقدير']
            ];

            this.studentsData.forEach(student => {
                exportData.push([
                    student.name,
                    student.notebook,
                    student.homework,
                    student.behavior,
                    student.participation,
                    student.continuous.toFixed(2),
                    student.test,
                    student.exam,
                    student.average.toFixed(2),
                    student.comment
                ]);
            });

            exportData.push([]);
            exportData.push(['الإحصائيات']);
            exportData.push(['عدد التلاميذ', this.studentsData.length]);
            exportData.push(['المعدل العام', (this.studentsData.reduce((sum, s) => sum + s.average, 0) / this.studentsData.length).toFixed(2)]);
            exportData.push(['أعلى معدل', Math.max(...this.studentsData.map(s => s.average)).toFixed(2)]);
            exportData.push(['أقل معدل', Math.min(...this.studentsData.map(s => s.average)).toFixed(2)]);
            exportData.push(['الناجحون', this.studentsData.filter(s => s.average >= 10).length]);
            exportData.push(['الراسبون', this.studentsData.filter(s => s.average < 10).length]);

            const wb = XLSX.utils.book_new();
            const ws = XLSX.utils.aoa_to_sheet(exportData);
            
            if (!ws['!cols']) ws['!cols'] = [];
            for (let i = 0; i < exportData[0].length; i++) {
                ws['!cols'][i] = { width: 15 };
            }

            XLSX.utils.book_append_sheet(wb, ws, "نقاط التلاميذ");
            
            const now = new Date();
            const sheetName = this.currentSheetName || 'غير محدد';
            const filename = `نقاط_التلاميذ_${sheetName}_${now.getFullYear()}-${(now.getMonth()+1).toString().padStart(2, '0')}-${now.getDate().toString().padStart(2, '0')}.xlsx`;
            
            XLSX.writeFile(wb, filename);
            
            this.hideLoading();
            this.showToast('تم تصدير البيانات بنجاح', 'success');
        } catch (error) {
            this.hideLoading();
            this.handleError(error, 'تصدير البيانات');
        }
    }

    // الإحصائيات والتقارير
    updateStatistics() {
        if (this.studentsData.length === 0) {
            this.resetStatisticsDisplay();
            this.showToast('لا توجد بيانات لعرض الإحصائيات', 'warning');
            return;
        }

        const totalStudents = this.studentsData.length;
        const averages = this.studentsData.map(s => s.average);
        const overallAverage = averages.reduce((sum, avg) => sum + avg, 0) / totalStudents;
        const highestGrade = Math.max(...averages);
        const lowestGrade = Math.min(...averages);
        const passedStudents = this.studentsData.filter(s => s.average >= 10).length;
        const failedStudents = totalStudents - passedStudents;

        document.getElementById('totalStudents').textContent = totalStudents;
        document.getElementById('overallAverage').textContent = overallAverage.toFixed(2);
        document.getElementById('highestGrade').textContent = highestGrade.toFixed(2);
        document.getElementById('lowestGrade').textContent = lowestGrade.toFixed(2);
        document.getElementById('passedStudents').textContent = passedStudents;
        document.getElementById('failedStudents').textContent = failedStudents;

        this.updateGradeDistribution();
        this.updateAdvancedStatistics(averages, overallAverage);
        this.showToast('تم تحديث الإحصائيات', 'success');
    }

    resetStatisticsDisplay() {
        const elements = [
            'totalStudents', 'overallAverage', 'highestGrade', 'lowestGrade',
            'passedStudents', 'failedStudents'
        ];
        
        elements.forEach(id => {
            const element = document.getElementById(id);
            if (element) element.textContent = '0';
        });
        
        const distributionElement = document.getElementById('gradeDistribution');
        if (distributionElement) {
            distributionElement.innerHTML = '<p style="text-align: center; color: #6c757d;">لا توجد بيانات لعرض التوزيع</p>';
        }
        
        const advancedStatsElement = document.getElementById('advancedStatistics');
        if (advancedStatsElement) {
            advancedStatsElement.innerHTML = '';
        }
    }

    updateGradeDistribution() {
        const distribution = {};
        this.studentsData.forEach(student => {
            const comment = student.comment;
            distribution[comment] = (distribution[comment] || 0) + 1;
        });

        const totalStudents = this.studentsData.length;
        let distributionHTML = '';
        
        Object.entries(distribution).forEach(([grade, count]) => {
            const percentage = ((count / totalStudents) * 100).toFixed(1);
            const percentageWidth = Math.max(10, percentage);
            
            distributionHTML += `
                <div class="distribution-item">
                    <div class="distribution-header">
                        <span class="grade-name">${grade}</span>
                        <span class="grade-count">${count} تلميذ (${percentage}%)</span>
                    </div>
                    <div class="distribution-bar">
                        <div class="distribution-fill" style="width: ${percentageWidth}%; background: ${this.getGradeColor(grade)};"></div>
                    </div>
                </div>
            `;
        });
        
        const distributionElement = document.getElementById('gradeDistribution');
        if (distributionElement) {
            distributionElement.innerHTML = distributionHTML || '<p style="text-align: center; color: #6c757d;">لا توجد بيانات لعرض التوزيع</p>';
        }
    }

    getGradeColor(gradeComment) {
        const colorMap = {
            'نتائج ممتازة ومرضية واصل': '#28a745',
            'نتائج جيدة و مرضية واصل': '#20c997',
            'واصل الاجتهاد و المثابرة': '#17a2b8',
            'نتائج مقبولة بإمكانك تحسينها': '#ffc107',
            'بمقدورك تحقيق نتائج أفضل': '#fd7e14',
            'ينقصك الحرص و التركيز': '#dc3545',
            'احذر التهاون': '#f10707ff'
        };
        
        return colorMap[gradeComment] || '#6c757d';
    }

    updateAdvancedStatistics(averages, overallAverage) {
        const variance = averages.reduce((sum, avg) => sum + Math.pow(avg - overallAverage, 2), 0) / averages.length;
        const standardDeviation = Math.sqrt(variance);
        
        const sortedAverages = [...averages].sort((a, b) => a - b);
        const median = sortedAverages.length % 2 === 0 
            ? (sortedAverages[sortedAverages.length / 2 - 1] + sortedAverages[sortedAverages.length / 2]) / 2
            : sortedAverages[Math.floor(sortedAverages.length / 2)];
        
        const frequency = {};
        let maxFrequency = 0;
        let mode = [];
        
        averages.forEach(avg => {
            const rounded = Math.round(avg * 10) / 10;
            frequency[rounded] = (frequency[rounded] || 0) + 1;
            
            if (frequency[rounded] > maxFrequency) {
                maxFrequency = frequency[rounded];
                mode = [rounded];
            } else if (frequency[rounded] === maxFrequency) {
                mode.push(rounded);
            }
        });
        
        let advancedStatsElement = document.getElementById('advancedStatistics');
        if (!advancedStatsElement) {
            advancedStatsElement = document.createElement('div');
            advancedStatsElement.id = 'advancedStatistics';
            advancedStatsElement.className = 'stat-card full-width';
            document.querySelector('#statistics-tab').appendChild(advancedStatsElement);
        }
        
        advancedStatsElement.innerHTML = `
            <h3>إحصائيات متقدمة</h3>
            <div class="advanced-stats-grid">
                <div class="advanced-stat">
                    <span class="stat-label">الوسيط (Median)</span>
                    <span class="stat-value">${median.toFixed(2)}</span>
                </div>
                <div class="advanced-stat">
                    <span class="stat-label">المنوال (Mode)</span>
                    <span class="stat-value">${mode.join(', ')}</span>
                </div>
                <div class="advanced-stat">
                    <span class="stat-label">الانحراف المعياري</span>
                    <span class="stat-value">${standardDeviation.toFixed(2)}</span>
                </div>
                <div class="advanced-stat">
                    <span class="stat-label">التباين (Variance)</span>
                    <span class="stat-value">${variance.toFixed(2)}</span>
                </div>
            </div>
        `;
    }

    exportStatistics() {
    let students = [];

    // الحالة الأولى: تصدير إحصائيات صفحة واحدة تمت معاينتها
    if (this.studentsData && this.studentsData.length > 0) {
        students = this.studentsData;
    } 
    // الحالة الثانية: تصدير الإحصائيات الشاملة لكل الصفحات المعالجة
    else if (this.allProcessedSheets && this.allProcessedSheets.size > 0) {
        this.allProcessedSheets.forEach(sheet => {
            students = students.concat(sheet.students);
        });
    }

    if (students.length === 0) {
        this.showToast('⚠️ لا توجد بيانات تلاميذ لتصدير الإحصائيات', 'warning');
        return;
    }

    // توليد التقرير
    const statisticsData = this.generateStatisticsReportFrom(students);
    const blob = new Blob([statisticsData], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');

    const now = new Date();
    const filename = `إحصائيات_${now.getFullYear()}-${(now.getMonth()+1).toString().padStart(2, '0')}-${now.getDate().toString().padStart(2, '0')}.txt`;

    a.href = url;
    a.download = filename;
    a.click();

    URL.revokeObjectURL(url);
    this.showToast('✅ تم تصدير التقرير الإحصائي بنجاح', 'success');
}
   generateStatisticsReportFrom(students) {
    const averages = students.map(s => s.average);
    const overallAverage = averages.reduce((sum, avg) => sum + avg, 0) / students.length;
    const passedStudents = students.filter(s => s.average >= 10).length;

    let report = `📊 تقرير إحصائي - ${new Date().toLocaleDateString('ar-DZ')}\n`;
    report += '='.repeat(60) + '\n\n';

    report += `📌 عدد التلاميذ: ${students.length}\n`;
    report += `📌 المعدل العام: ${overallAverage.toFixed(2)}\n`;
    report += `📌 أعلى معدل: ${Math.max(...averages).toFixed(2)}\n`;
    report += `📌 أقل معدل: ${Math.min(...averages).toFixed(2)}\n`;
    report += `📌 عدد الناجحين: ${passedStudents}\n`;
    report += `📌 عدد الراسبين: ${students.length - passedStudents}\n`;
    report += `📌 نسبة النجاح: ${((passedStudents / students.length) * 100).toFixed(1)}%\n\n`;

    report += '📈 توزيع التقديرات:\n';
    report += '-'.repeat(40) + '\n';

    const distribution = {};
    students.forEach(student => {
        distribution[student.comment] = (distribution[student.comment] || 0) + 1;
    });

    Object.entries(distribution).forEach(([grade, count]) => {
        const percentage = ((count / students.length) * 100).toFixed(1);
        report += `${grade}: ${count} تلميذ (${percentage}%)\n`;
    });

    return report;
}
    // الإعدادات
    loadGradeSettings() {
        const savedSettings = localStorage.getItem('gradeSettings');
        if (savedSettings) {
            try {
                this.gradeSettings = JSON.parse(savedSettings);
                this.showToast('تم تحميل الإعدادات المحفوظة', 'success');
            } catch (error) {
                console.error('❌ خطأ في تحميل إعدادات التقديرات:', error);
                this.showToast('خطأ في تحميل الإعدادات، سيتم استخدام الإعدادات الافتراضية', 'error');
            }
        }
        this.updateGradeSettingsTable();
    }

    saveGradeSettings() {
        const inputs = document.querySelectorAll('#gradeSettingsBody input');
        const newSettings = [];
        
        try {
            for (let i = 0; i < inputs.length; i += 2) {
                const minInput = inputs[i];
                const commentInput = inputs[i + 1];
                
                const minValue = parseFloat(minInput.value);
                if (isNaN(minValue) || minValue < 0 || minValue > 20) {
                    throw new Error(`القيمة ${minInput.value} غير صالحة للحد الأدنى`);
                }
                
                newSettings.push({
                    min: minValue,
                    max: i === 0 ? 20 : parseFloat(inputs[i - 2].value) - 0.01,
                    comment: commentInput.value.trim() || 'غير محدد'
                });
            }
            
            newSettings.sort((a, b) => b.min - a.min);
            
            for (let i = 0; i < newSettings.length; i++) {
                if (i === 0) {
                    newSettings[i].max = 20;
                } else {
                    newSettings[i].max = newSettings[i - 1].min - 0.01;
                }
            }
            
            this.gradeSettings = newSettings;
            localStorage.setItem('gradeSettings', JSON.stringify(this.gradeSettings));
            this.updateStudentsWithNewSettings();
            this.showToast('تم حفظ إعدادات التقديرات بنجاح', 'success');
        } catch (error) {
            this.showToast(`خطأ في الحفظ: ${error.message}`, 'error');
        }
    }

    resetGradeSettings() {
        if (confirm('هل تريد استعادة الإعدادات الافتراضية؟ سيتم فقدان التغييرات الحالية.')) {
            this.gradeSettings = [
                { min: 18, max: 20, comment: 'نتائج ممتازة ومرضية واصل' },
                { min: 16, max: 17.99, comment: 'نتائج جيدة و مرضية واصل' },
                { min: 14, max: 15.99, comment: 'واصل الاجتهاد و المثابرة' },
                { min: 12, max: 13.99, comment: 'نتائج مقبولة بإمكانك تحسينها' },
                { min: 10, max: 11.99, comment: 'بمقدورك تحقيق نتائج أفضل' },
                { min: 7, max: 9.99, comment: 'ينقصك الحرص و التركيز' },
                { min: 0, max: 6.99, comment: 'احذر التهاون' }
            ];
            localStorage.removeItem('gradeSettings');
            this.updateGradeSettingsTable();
            this.updateStudentsWithNewSettings();
            this.showToast('تم استعادة الإعدادات الافتراضية بنجاح', 'success');
        }
    }

    updateGradeSettingsTable() {
        const tbody = document.getElementById('gradeSettingsBody');
        if (!tbody) return;
        
        tbody.innerHTML = '';
        
        const sortedSettings = [...this.gradeSettings].sort((a, b) => b.min - a.min);
        
        sortedSettings.forEach((setting, index) => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${index + 1}</td>
                <td>
                    <input type="number" 
                           value="${setting.min}" 
                           min="0" 
                           max="20" 
                           step="0.5"
                           onchange="app.validateGradeSetting(this, ${index})">
                </td>
                <td>${setting.max.toFixed(2)}</td>
                <td>
                    <input type="text" 
                           value="${setting.comment}" 
                           style="width: 100%;"
                           placeholder="أدخل العبارة التقديرية">
                </td>
                <td>
                    <button class="btn btn-danger btn-sm" onclick="app.removeGradeSetting(${index})">حذف</button>
                </td>
            `;
            tbody.appendChild(tr);
        });
        
        const addRow = document.createElement('tr');
        addRow.innerHTML = `
            <td colspan="5" style="text-align: center;">
                <button class="btn btn-primary" onclick="app.addNewGradeSetting()">
                    + إضافة تقدير جديد
                </button>
            </td>
        `;
        tbody.appendChild(addRow);
    }

    validateGradeSetting(input, index) {
        const value = parseFloat(input.value);
        if (isNaN(value) || value < 0 || value > 20) {
            input.style.borderColor = 'red';
            this.showToast('القيمة يجب أن تكون بين 0 و 20', 'error');
            return false;
        }
        
        const sortedSettings = [...this.gradeSettings].sort((a, b) => b.min - a.min);
        for (let i = 0; i < sortedSettings.length; i++) {
            if (i !== index && Math.abs(sortedSettings[i].min - value) < 0.5) {
                input.style.borderColor = 'orange';
                this.showToast('تنبيه: قد يكون هناك تداخل في التقديرات', 'warning');
                return true;
            }
        }
        
        input.style.borderColor = '';
        return true;
    }

    addNewGradeSetting() {
        const minValues = this.gradeSettings.map(s => s.min);
        const newMin = Math.min(...minValues) - 1;
        
        if (newMin < 0) {
            this.showToast('لا يمكن إضافة المزيد من التقديرات', 'error');
            return;
        }
        
        this.gradeSettings.push({
            min: newMin,
            max: newMin + 0.99,
            comment: 'تقدير جديد'
        });
        
        this.updateGradeSettingsTable();
        this.showToast('تم إضافة تقدير جديد', 'success');
    }

    removeGradeSetting(index) {
        if (this.gradeSettings.length <= 1) {
            this.showToast('يجب أن يكون هناك تقدير واحد على الأقل', 'error');
            return;
        }
        
        if (confirm('هل تريد حذف هذا التقدير؟')) {
            this.gradeSettings.splice(index, 1);
            this.updateGradeSettingsTable();
            this.updateStudentsWithNewSettings();
            this.showToast('تم حذف التقدير', 'success');
        }
    }

    updateStudentsWithNewSettings() {
        if (this.studentsData.length > 0) {
            this.studentsData.forEach((student, index) => {
                student.comment = this.getGradeComment(student.average);
            });
            this.updateGradesTable();
            this.showToast('تم تحديث تقديرات التلاميذ', 'success');
        }
    }

    exportSettings() {
        const settingsData = {
            gradeSettings: this.gradeSettings,
            exportDate: new Date().toISOString(),
            version: '1.0'
        };
        
        const dataStr = JSON.stringify(settingsData, null, 2);
        const blob = new Blob([dataStr], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        
        a.href = url;
        a.download = `إعدادات_التقديرات_${new Date().getTime()}.json`;
        a.click();
        
        URL.revokeObjectURL(url);
        this.showToast('تم تصدير الإعدادات', 'success');
    }

    importSettings(event) {
        const file = event.target.files[0];
        if (!file) return;
        
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const settingsData = JSON.parse(e.target.result);
                
                if (!settingsData.gradeSettings || !Array.isArray(settingsData.gradeSettings)) {
                    throw new Error('تنسيق الملف غير صالح');
                }
                
                this.gradeSettings = settingsData.gradeSettings;
                localStorage.setItem('gradeSettings', JSON.stringify(this.gradeSettings));
                this.updateGradeSettingsTable();
                this.updateStudentsWithNewSettings();
                
                this.showToast('تم استيراد الإعدادات بنجاح', 'success');
            } catch (error) {
                this.showToast(`خطأ في استيراد الإعدادات: ${error.message}`, 'error');
            }
        };
        
        reader.readAsText(file);
        event.target.value = '';
    }

    resetAllData() {
        if (confirm('هل أنت متأكد من مسح جميع البيانات؟ هذا الإجراء لا يمكن التراجع عنه.')) {
            if (confirm('⚠️ تأكيد نهائي: جميع البيانات بما في ذلك التلاميذ والنقاط والإعدادات سيتم مسحها.')) {
                this.studentsData = [];
                this.selectedRows.clear();
                this.currentSheet = null;
                this.allProcessedSheets.clear();
                this.currentWorkbook = null;
                
                localStorage.removeItem('gradeSettings');
                
                this.updateGradesTable();
                this.updateStatistics();
                this.loadGradeSettings();
                
                document.getElementById('previewSection').style.display = 'none';
                document.getElementById('sheetSelector').style.display = 'none';
                document.getElementById('bulkProcessingSection').style.display = 'none';
                document.getElementById('bulkProcessingResults').style.display = 'none';
                document.getElementById('exportAllSheets').style.display = 'none';
                document.getElementById('previewAllSheets').style.display = 'none';
                
                this.showToast('تم مسح جميع البيانات بنجاح', 'success');
            }
        }
    }

    // =============================================================================
    // معالجة جميع صفحات Excel - المحسنة
    // =============================================================================

    // معالجة جميع صفحات الملف
    processAllSheets() {
        if (!this.currentWorkbook) {
            this.showToast('لم يتم تحميل أي ملف', 'warning');
            return;
        }

        this.showLoading();
        this.allProcessedSheets.clear();

        try {
            const sheetNames = this.currentWorkbook.SheetNames;
            let totalStudents = 0;
            let processedCount = 0;
            let totalPassedStudents = 0; // ✅ إضافة عداد التلاميذ الناجحين
            // معالجة كل صفحة
            sheetNames.forEach(sheetName => {
                try {
                    const worksheet = this.currentWorkbook.Sheets[sheetName];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, {header: 1, defval: ""});
                    
                    const sheetResult = this.processSingleSheet(jsonData, sheetName);
                    
                    if (sheetResult.students.length > 0) {
                        this.allProcessedSheets.set(sheetName, sheetResult);
                        totalStudents += sheetResult.students.length;
                        totalPassedStudents += sheetResult.stats.passedStudents; // ✅ جمع الناجحين
                        processedCount++;
                    }
                } catch (error) {
                    console.warn(`خطأ في معالجة الصفحة ${sheetName}:`, error);
                }
            });

            this.hideLoading();
            
            if (processedCount === 0) {
                this.showToast('لم يتم العثور على بيانات صالحة في أي صفحة', 'warning');
                return;
            }

            this.showBulkProcessingResults(processedCount, totalStudents, sheetNames.length, totalPassedStudents);
            this.showToast(`تم معالجة ${processedCount} صفحة بنجاح - إجمالي ${totalStudents} تلميذ`, 'success');
            
        } catch (error) {
            this.hideLoading();
            this.handleError(error, 'معالجة جميع الصفحات');
        }
    }

    // معالجة صفحة واحدة
    processSingleSheet(data, sheetName) {
        // اكتشاف الأعمدة مع التحسينات الجديدة
        const detectionResult = this.detectColumns(data);
        
        // استخراج الطلاب باستخدام التحسينات الجديدة
        const students = this.extractStudentsFromData(data, detectionResult);
        
        // حساب النقاط والتقديرات
        students.forEach((student) => {
            this.calculateStudentGrades(student);
        });

        // حساب إحصائيات الصفحة
        const stats = this.calculateSheetStatistics(students);

        return {
            sheetName,
            students,
            columns: detectionResult.columns,
            detectionResult,
            stats,
            processedAt: new Date().toISOString()
        };
    }

    // اكتشاف الأعمدة في البيانات - المحسنة
    detectColumns(data) {
        const columns = {
            lastNameColumn: -1,
            firstNameColumn: -1,
            notebookColumn: -1,
            homeworkColumn: -1,
            behaviorColumn: -1,
            participationColumn: -1,
            continuousColumn: -1,
            testColumn: -1,
            examColumn: -1
        };

        const keywords = {
            lastNameColumn: ['اللقب', 'لقب', 'اسم العائلة', 'Nom', 'Last Name', 'Surname'],
            firstNameColumn: ['الاسم', 'اسم','إسم','الإسم', 'الاسم الشخصي', 'Prénom', 'First Name', 'Name'],
            notebookColumn: ['الكراس', 'كراس', 'Cahier', 'Notebook'],
            homeworkColumn: ['الواجبات', 'واجبات', 'Devoirs', 'Homework'],
            behaviorColumn: ['السلوك', 'سلوك', 'Comportement', 'Behavior'],
            participationColumn: ['المشاركة', 'مشاركة', 'Participation'],
            continuousColumn: ['التقويم المستمر', 'تقويم مستمر','معدل تقويم النشاطات','التقويم', 'تقويم','Continu', 'Continuous'],
            testColumn: ['الفرض', 'فرض','وقفة تقييمية','وقفة تقويمية', 'Devoir', 'Test'],
            examColumn: ['الاختبار', 'اختبار','الإختبار','إختبار','الامتحان' , 'امتحان','Examen', 'Exam']
        };

        let headerRowIndex = -1;
        let studentRowsStart = -1;
        let studentRowsEnd = -1;

        // البحث عن صف العناوين الذي يحتوي على اللقب والاسم معاً
        for (let rowIndex = 0; rowIndex < Math.min(20, data.length); rowIndex++) {
            const row = data[rowIndex];
            let foundLastName = false;
            let foundFirstName = false;
            
            for (let colIndex = 0; colIndex < row.length; colIndex++) {
                const cellValue = String(row[colIndex] || '').trim();
                
                // البحث عن عمود اللقب
                if (!foundLastName && keywords.lastNameColumn.some(keyword => 
                    cellValue.toLowerCase().includes(keyword.toLowerCase()))) {
                    columns.lastNameColumn = colIndex;
                    foundLastName = true;
                }
                
                // البحث عن عمود الاسم
                if (!foundFirstName && keywords.firstNameColumn.some(keyword => 
                    cellValue.toLowerCase().includes(keyword.toLowerCase()))) {
                    columns.firstNameColumn = colIndex;
                    foundFirstName = true;
                }
                
                // البحث عن الأعمدة الأخرى
                for (const [column, keywordList] of Object.entries(keywords)) {
                    if (columns[column] === -1 && column !== 'lastNameColumn' && column !== 'firstNameColumn') {
                        if (keywordList.some(keyword => 
                            cellValue.toLowerCase().includes(keyword.toLowerCase()))) {
                            columns[column] = colIndex;
                        }
                    }
                }
            }
            
            // إذا وجدنا كلًا من اللقب والاسم في نفس الصف، فهذا هو صف العناوين
            if (foundLastName && foundFirstName) {
                headerRowIndex = rowIndex;
                break;
            }
        }

        // إذا لم نجد صف العناوين الذي يحتوي على اللقب والاسم معاً، نبحث عن أي صف يحتوي على أي منهما
        if (headerRowIndex === -1) {
            for (let rowIndex = 0; rowIndex < Math.min(20, data.length); rowIndex++) {
                const row = data[rowIndex];
                
                for (let colIndex = 0; colIndex < row.length; colIndex++) {
                    const cellValue = String(row[colIndex] || '').trim();
                    
                    for (const [column, keywordList] of Object.entries(keywords)) {
                        if (columns[column] === -1 && keywordList.some(keyword => 
                            cellValue.toLowerCase().includes(keyword.toLowerCase()))) {
                            columns[column] = colIndex;
                            headerRowIndex = rowIndex;
                        }
                    }
                }
                
                if (headerRowIndex !== -1) break;
            }
        }

        // تحديد بداية ونهاية صفوف التلاميذ
        if (headerRowIndex !== -1) {
            studentRowsStart = headerRowIndex + 1;
            studentRowsEnd = data.length - 1;
            
            // البحث عن نهاية صفوف التلاميذ
            for (let i = studentRowsStart; i < data.length; i++) {
                const row = data[i];
                const hasStudentData = row && 
                    (columns.lastNameColumn !== -1 && row[columns.lastNameColumn] && String(row[columns.lastNameColumn]).trim()) &&
                    (columns.firstNameColumn !== -1 && row[columns.firstNameColumn] && String(row[columns.firstNameColumn]).trim());
                
                if (!hasStudentData) {
                    studentRowsEnd = i - 1;
                    break;
                }
            }
        }

        return {
            columns,
            headerRowIndex,
            studentRowsStart,
            studentRowsEnd
        };
    }

    // استخراج الطلاب من البيانات - المحسنة
    extractStudentsFromData(data, detectionResult) {
        const { columns, studentRowsStart, studentRowsEnd } = detectionResult;
        const students = [];
        const { lastNameColumn, firstNameColumn } = columns;
        
        if (lastNameColumn === -1 || firstNameColumn === -1 || studentRowsStart === -1) {
            console.warn('⚠️ لم يتم العثور على بيانات تلاميذ كافية');
            return students;
        }

        // استخراج بيانات الطلاب من الصفوف المحددة
        for (let i = studentRowsStart; i <= studentRowsEnd; i++) {
            const row = data[i];
            if (!row) continue;
            
            const lastName = String(row[lastNameColumn] || '').trim();
            const firstName = String(row[firstNameColumn] || '').trim();
            
            // التأكد من وجود بيانات الاسم واللقب
            if (lastName && firstName && lastName !== '' && firstName !== '') {
                const student = {
                    name: `${lastName} ${firstName}`,
                    lastName,
                    firstName,
                    notebook: this.parseGrade(row[columns.notebookColumn]),
                    homework: this.parseGrade(row[columns.homeworkColumn]),
                    behavior: this.parseGrade(row[columns.behaviorColumn]),
                    participation: this.parseGrade(row[columns.participationColumn]),
                    test: this.parseGrade(row[columns.testColumn]),
                    exam: this.parseGrade(row[columns.examColumn]),
                    continuous: 0,
                    average: 0,
                    comment: '',
                    originalRow: i + 1, // حفظ رقم الصف الأصلي
                    originalData: row
                };

                // إذا كان هناك تقويم مستمر لكن لا توجد علامات جزئية
                const continuous = this.parseGrade(row[columns.continuousColumn]);
                if (continuous > 0 && student.notebook === 0 && student.homework === 0 && 
                    student.behavior === 0 && student.participation === 0) {
                    student.notebook = student.homework = student.behavior = student.participation = continuous;
                }

                students.push(student);
            }
        }

        console.log(`✅ تم استخراج ${students.length} تلميذ من الصفوف ${studentRowsStart + 1} إلى ${studentRowsEnd + 1}`);
        return students;
    }

    // دالة مساعدة لتحويل القيم إلى أرقام
    parseGrade(value) {
        if (value === null || value === undefined || value === '') return 0;
        
        // تحويل النصوص إلى أرقام
        const numValue = parseFloat(value);
        return isNaN(numValue) ? 0 : Math.max(0, Math.min(20, numValue));
    }

    // حساب نقاط وتقديرات الطالب
    calculateStudentGrades(student) {
        student.continuous = (student.notebook + student.homework + student.behavior + student.participation) / 4;
        student.average = (student.continuous + student.test + (student.exam * 3)) / 5;
        student.comment = this.getGradeComment(student.average);
    }

    // حساب إحصائيات الصفحة
    calculateSheetStatistics(students) {
        if (students.length === 0) {
            return {
                totalStudents: 0,
                overallAverage: 0,
                highestGrade: 0,
                lowestGrade: 0,
                passedStudents: 0,
                failedStudents: 0
            };
        }

        const averages = students.map(s => s.average);
        const overallAverage = averages.reduce((sum, avg) => sum + avg, 0) / students.length;
        const passedStudents = students.filter(s => s.average >= 10).length;

        return {
            totalStudents: students.length,
            overallAverage: parseFloat(overallAverage.toFixed(2)),
            highestGrade: parseFloat(Math.max(...averages).toFixed(2)),
            lowestGrade: parseFloat(Math.min(...averages).toFixed(2)),
            passedStudents,
            failedStudents: students.length - passedStudents
        };
    }

    // عرض نتائج المعالجة الشاملة
    showBulkProcessingResults(processedCount, totalStudents, totalSheets,passedStudents) {
        const resultsDiv = document.getElementById('bulkProcessingResults');
        const summaryDiv = document.getElementById('processingSummary');
        const sheetsListDiv = document.getElementById('processedSheetsList');
        const successRate = totalStudents > 0 ? Math.round((passedStudents / totalStudents) * 100) : 0;

        // تحديث الملخص
        summaryDiv.innerHTML = `
            <div class="summary-item">
                <span class="summary-number">${processedCount}</span>
                <span class="summary-label">صفحة معالجة</span>
            </div>
            <div class="summary-item">
                <span class="summary-number">${totalStudents}</span>
                <span class="summary-label">إجمالي التلاميذ</span>
            </div>
            <div class="summary-item">
                <span class="summary-number">${totalSheets}</span>
                <span class="summary-label">إجمالي الصفحات</span>
            </div>
            <div class="summary-item">
                <span class="summary-number">${successRate}%</span>
                <span class="summary-label">نسبة النجاح</span>
            </div>
        `;

        // عرض قائمة الصفحات المعالجة
        sheetsListDiv.innerHTML = '';
        this.allProcessedSheets.forEach((sheetData, sheetName) => {
            const sheetCard = this.createSheetCard(sheetData);
            sheetsListDiv.appendChild(sheetCard);
        });

        // إظهار النتائج والأزرار
        resultsDiv.style.display = 'block';
        document.getElementById('exportAllSheets').style.display = 'inline-block';
        document.getElementById('previewAllSheets').style.display = 'inline-block';
    }

    // إنشاء بطاقة عرض للصفحة
    createSheetCard(sheetData) {
        const card = document.createElement('div');
        card.className = 'sheet-card';
        
        const { sheetName, students, stats } = sheetData;
        
        card.innerHTML = `
            <div class="sheet-card-header">
                <span class="sheet-name">${sheetName}</span>
            </div>
            <div class="sheet-stats">
                <div class="sheet-stat">
                    <span class="sheet-stat-number">${stats.totalStudents}</span>
                    <span class="sheet-stat-label">تلميذ</span>
                </div>
                <div class="sheet-stat">
                    <span class="sheet-stat-number">${stats.overallAverage}</span>
                    <span class="sheet-stat-label">المعدل</span>
                </div>
                <div class="sheet-stat">
                    <span class="sheet-stat-number">${stats.passedStudents}</span>
                    <span class="sheet-stat-label">ناجح</span>
                </div>
            </div>
            <div class="sheet-actions">
                <button class="sheet-action-btn sheet-action-preview" onclick="app.previewSheet('${sheetName}')">
                    معاينة
                </button>
                <button class="sheet-action-btn sheet-action-export" onclick="app.exportSingleSheet('${sheetName}')">
                    تصدير
                </button>
            </div>
        `;
        
        return card;
    }

    // معاينة صفحة محددة
    previewSheet(sheetName) {
        const sheetData = this.allProcessedSheets.get(sheetName);
        if (!sheetData) {
            this.showToast('لم يتم العثور على بيانات الصفحة', 'error');
            return;
        }

        // حفظ البيانات مؤقتاً وعرضها في تبويب النقاط
        this.studentsData = sheetData.students;
        this.currentSheetName = sheetName;
        this.updateGradesTable();
        this.showTab('grades');
        
        this.showToast(`تم تحميل بيانات الصفحة: ${sheetName}`, 'success');
    }

    // تصدير صفحة محددة
    exportSingleSheet(sheetName) {
        const sheetData = this.allProcessedSheets.get(sheetName);
        if (!sheetData) {
            this.showToast('لم يتم العثور على بيانات الصفحة', 'error');
            return;
        }

        this.exportSheetToExcel(sheetData, `${sheetName}_معالجة`);
    }

    // تصدير جميع الصفحات المعالجة
    exportAllProcessedSheets() {
        if (this.allProcessedSheets.size === 0) {
            this.showToast('لا توجد صفحات معالجة للتصدير', 'warning');
            return;
        }

        this.showLoading();

        try {
            const wb = XLSX.utils.book_new();
            let totalStats = {
                totalStudents: 0,
                totalSheets: 0,
                overallAverages: []
            };

            // إنشاء صفحة لكل مجموعة طلاب معالجة
            this.allProcessedSheets.forEach((sheetData, sheetName) => {
                const exportData = this.prepareSheetExportData(sheetData);
                const ws = XLSX.utils.aoa_to_sheet(exportData);
                XLSX.utils.book_append_sheet(wb, ws, this.truncateSheetName(sheetName));
                
                // جمع الإحصائيات
                totalStats.totalStudents += sheetData.stats.totalStudents;
                totalStats.totalSheets++;
                totalStats.overallAverages.push(sheetData.stats.overallAverage);
            });

            // إضافة صفحة الإحصائيات العامة
            const statsData = this.prepareGlobalStatsData(totalStats);
            const wsStats = XLSX.utils.aoa_to_sheet(statsData);
            XLSX.utils.book_append_sheet(wb, wsStats, "الإحصائيات_العامة");

            // إضافة صفحة الملخص
            const summaryData = this.prepareSummaryData();
            const wsSummary = XLSX.utils.aoa_to_sheet(summaryData);
            XLSX.utils.book_append_sheet(wb, wsSummary, "ملخص_المعالجة");

            // حفظ الملف
            const now = new Date();
            const filename = `الملف_الشامل_${now.getFullYear()}-${(now.getMonth()+1).toString().padStart(2, '0')}-${now.getDate().toString().padStart(2, '0')}.xlsx`;
            
            XLSX.writeFile(wb, filename);
            
            this.hideLoading();
            this.showToast(`تم تصدير ${this.allProcessedSheets.size} صفحة بنجاح`, 'success');
            
        } catch (error) {
            this.hideLoading();
            this.handleError(error, 'تصدير جميع الصفحات');
        }
    }

    // إعداد بيانات التصدير للصفحة
    prepareSheetExportData(sheetData) {
        const { students, stats, sheetName } = sheetData;
        
        const exportData = [
            [`تقرير الصفحة: ${sheetName}`],
            [''],
            ['اللقب والاسم', 'الكراس', 'الواجبات', 'السلوك', 'المشاركة', 'التقويم المستمر', 'الفرض', 'الاختبار', 'المعدل', 'التقدير']
        ];

        // بيانات الطلاب
        students.forEach(student => {
            exportData.push([
                student.name,
                student.notebook,
                student.homework,
                student.behavior,
                student.participation,
                student.continuous.toFixed(2),
                student.test,
                student.exam,
                student.average.toFixed(2),
                student.comment
            ]);
        });

        // إحصائيات الصفحة
        exportData.push(['']);
        exportData.push(['إحصائيات الصفحة:']);
        exportData.push(['عدد التلاميذ', stats.totalStudents]);
        exportData.push(['المعدل العام', stats.overallAverage]);
        exportData.push(['أعلى معدل', stats.highestGrade]);
        exportData.push(['أقل معدل', stats.lowestGrade]);
        exportData.push(['الناجحون', stats.passedStudents]);
        exportData.push(['الراسبون', stats.failedStudents]);
        exportData.push(['نسبة النجاح', `${((stats.passedStudents / stats.totalStudents) * 100).toFixed(1)}%`]);

        return exportData;
    }

    // إعداد بيانات الإحصائيات العامة
    prepareGlobalStatsData(totalStats) {
        const overallAverage = totalStats.overallAverages.reduce((sum, avg) => sum + avg, 0) / totalStats.overallAverages.length;
        
        return [
            ['الإحصائيات العامة للملف'],
            [''],
            ['إجمالي عدد الصفحات المعالجة', totalStats.totalSheets],
            ['إجمالي عدد التلاميذ', totalStats.totalStudents],
            ['المعدل العام للجميع', overallAverage.toFixed(2)],
            ['متوسط التلاميذ لكل صفحة', Math.round(totalStats.totalStudents / totalStats.totalSheets)],
            [''],
            ['تفاصيل الصفحات:'],
            ['اسم الصفحة', 'عدد التلاميذ', 'المعدل', 'الناجحون', 'نسبة النجاح'],
            ...Array.from(this.allProcessedSheets.entries()).map(([name, data]) => [
                name,
                data.stats.totalStudents,
                data.stats.overallAverage,
                data.stats.passedStudents,
                `${((data.stats.passedStudents / data.stats.totalStudents) * 100).toFixed(1)}%`
            ])
        ];
    }

    // إعداد بيانات الملخص
    prepareSummaryData() {
        return [
            ['ملخص عملية المعالجة'],
            [''],
            ['تاريخ المعالجة', new Date().toLocaleDateString('ar-DZ')],
            ['وقت المعالجة', new Date().toLocaleTimeString('ar-DZ')],
            ['عدد الصفحات المعالجة', this.allProcessedSheets.size],
            ['إجمالي التلاميذ', Array.from(this.allProcessedSheets.values()).reduce((sum, sheet) => sum + sheet.stats.totalStudents, 0)],
            [''],
            ['الصفحات المعالجة:'],
            ...Array.from(this.allProcessedSheets.keys()).map(name => [name])
        ];
    }

    // تقصير اسم الصفحة إذا كان طويلاً
    truncateSheetName(name, maxLength = 25) {
        if (name.length <= maxLength) return name;
        return name.substring(0, maxLength - 3) + '...';
    }

    // تصدير صفحة مفردة
    exportSheetToExcel(sheetData, filenameSuffix = '') {
        const exportData = this.prepareSheetExportData(sheetData);
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(exportData);
        
        XLSX.utils.book_append_sheet(wb, ws, "النتائج");
        
        const now = new Date();
        const filename = `نتائج_${filenameSuffix}_${now.getFullYear()}-${(now.getMonth()+1).toString().padStart(2, '0')}-${now.getDate().toString().padStart(2, '0')}.xlsx`;
        
        XLSX.writeFile(wb, filename);
        this.showToast('تم تصدير الصفحة بنجاح', 'success');
    }

    // معاينة جميع الصفحات (عرض إحصائية عامة)
    previewAllSheets() {
        if (this.allProcessedSheets.size === 0) {
            this.showToast('لا توجد صفحات معالجة', 'warning');
            return;
        }

        this.showTab('statistics');
        this.updateGlobalStatistics();
    }

    // تحديث الإحصائيات العامة
    updateGlobalStatistics() {
        if (this.allProcessedSheets.size === 0) return;

        let totalStudents = 0;
        let totalAverages = 0;
        let allAverages = [];
        let totalPassed = 0;

        this.allProcessedSheets.forEach(sheetData => {
            totalStudents += sheetData.stats.totalStudents;
            totalAverages += sheetData.stats.overallAverage * sheetData.stats.totalStudents;
            totalPassed += sheetData.stats.passedStudents;
            allAverages = allAverages.concat(sheetData.students.map(s => s.average));
        });

        const overallAverage = totalAverages / totalStudents;
        const overallPassRate = (totalPassed / totalStudents) * 100;

        // تحديث واجهة الإحصائيات
        document.getElementById('totalStudents').textContent = totalStudents;
        document.getElementById('overallAverage').textContent = overallAverage.toFixed(2);
        document.getElementById('highestGrade').textContent = Math.max(...allAverages).toFixed(2);
        document.getElementById('lowestGrade').textContent = Math.min(...allAverages).toFixed(2);
        document.getElementById('passedStudents').textContent = totalPassed;
        document.getElementById('failedStudents').textContent = totalStudents - totalPassed;

        // تحديث توزيع التقديرات لجميع الطلاب
        this.updateGlobalGradeDistribution(allAverages);
    }

    // تحديث توزيع التقديرات لجميع الطلاب
    updateGlobalGradeDistribution(allAverages) {
        const distribution = {};
        
        allAverages.forEach(average => {
            const comment = this.getGradeComment(average);
            distribution[comment] = (distribution[comment] || 0) + 1;
        });

        const totalStudents = allAverages.length;
        let distributionHTML = '';
        
        Object.entries(distribution).forEach(([grade, count]) => {
            const percentage = ((count / totalStudents) * 100).toFixed(1);
            const percentageWidth = Math.max(10, percentage);
            
            distributionHTML += `
                <div class="distribution-item">
                    <div class="distribution-header">
                        <span class="grade-name">${grade}</span>
                        <span class="grade-count">${count} تلميذ (${percentage}%)</span>
                    </div>
                    <div class="distribution-bar">
                        <div class="distribution-fill" style="width: ${percentageWidth}%; background: ${this.getGradeColor(grade)};"></div>
                    </div>
                </div>
            `;
        });
        
        const distributionElement = document.getElementById('gradeDistribution');
        if (distributionElement) {
            distributionElement.innerHTML = distributionHTML;
        }
    }

    finalizeInit() {
        this.updateGradeInputs();
        console.log('✅ تم تهيئة التطبيق بالكامل بنجاح');
    }
    exportGlobalStatisticsAsText() {
    if (this.allProcessedSheets.size === 0) {
        this.showToast('لا توجد صفحات معالجة لتصدير الإحصائيات', 'warning');
        return;
    }

    let totalStudents = 0;
    let totalAverages = 0;
    let allAverages = [];
    let totalPassed = 0;

    this.allProcessedSheets.forEach(sheetData => {
        totalStudents += sheetData.stats.totalStudents;
        totalAverages += sheetData.stats.overallAverage * sheetData.stats.totalStudents;
        totalPassed += sheetData.stats.passedStudents;
        allAverages = allAverages.concat(sheetData.students.map(s => s.average));
    });

    const overallAverage = (totalAverages / totalStudents).toFixed(2);
    const highestGrade = Math.max(...allAverages).toFixed(2);
    const lowestGrade = Math.min(...allAverages).toFixed(2);
    const passRate = ((totalPassed / totalStudents) * 100).toFixed(1);
    const failedStudents = totalStudents - totalPassed;

    let text = '';
    text += `📊 الإحصائيات العامة:\n\n`;
    text += `- إجمالي التلاميذ: ${totalStudents}\n`;
    text += `- المعدل العام: ${overallAverage}\n`;
    text += `- أعلى معدل: ${highestGrade}\n`;
    text += `- أقل معدل: ${lowestGrade}\n`;
    text += `- عدد الناجحين: ${totalPassed}\n`;
    text += `- عدد الراسبين: ${failedStudents}\n`;
    text += `- نسبة النجاح: ${passRate}%\n\n`;

    text += `🔢 توزيع التقديرات:\n`;

    // حساب توزيع التقديرات
    const distribution = {};
    allAverages.forEach(avg => {
        const comment = this.getGradeComment(avg);
        distribution[comment] = (distribution[comment] || 0) + 1;
    });

    Object.entries(distribution).forEach(([grade, count]) => {
        const percentage = ((count / totalStudents) * 100).toFixed(1);
        text += `- ${grade}: ${count} تلميذ (${percentage}%)\n`;
    });

    // حفظ كنص
    const blob = new Blob([text], { type: 'text/plain;charset=utf-8' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);

    const now = new Date();
    const filename = `احصائيات_المعالجة_${now.getFullYear()}-${(now.getMonth()+1).toString().padStart(2, '0')}-${now.getDate().toString().padStart(2, '0')}.txt`;

    link.download = filename;
    link.click();

    URL.revokeObjectURL(link.href);
    this.showToast('تم تصدير الإحصائيات العامة كملف نصي', 'success');
}

}

// تهيئة التطبيق عند تحميل الصفحة
document.addEventListener('DOMContentLoaded', function() {
    window.app = new StudentGradesApp();
});