/**
 * מערכת ניהול הזמנות - עיבוד Excel
 * כולל חישובי אופטימיזציה, סינונים ודוחות
 */

class OrderManagementSystem {
    constructor() {
        this.data = [];
        this.filteredData = [];
        this.filters = {};
        this.allowOverage = false;
        
        this.initializeEventListeners();
        this.loadSettings();
    }

    /**
     * אתחול מאזינים לאירועים
     */
    initializeEventListeners() {
        // File upload
        document.getElementById('uploadBtn').addEventListener('click', () => {
            document.getElementById('excelFile').click();
        });
        
        document.getElementById('excelFile').addEventListener('change', (e) => {
            this.handleFileUpload(e.target.files[0]);
        });

        // Create sample file button
        document.getElementById('createSampleBtn').addEventListener('click', () => {
            this.createSampleFile();
        });

        // Search
        document.getElementById('searchInput').addEventListener('input', (e) => {
            this.handleSearch(e.target.value);
        });

        // Toggle settings
        document.getElementById('allowOverageToggle').addEventListener('change', (e) => {
            this.allowOverage = e.target.checked;
            this.saveSettings();
            this.recalculateOptimizations();
        });

        // Filters
        ['branchFilter', 'cityFilter', 'customerTypeFilter', 'categoryFilter'].forEach(id => {
            document.getElementById(id).addEventListener('change', () => {
                this.applyFilters();
            });
        });

        document.getElementById('clearFiltersBtn').addEventListener('click', () => {
            this.clearFilters();
        });

        // Export buttons
        document.getElementById('exportExcelBtn').addEventListener('click', () => {
            this.exportToExcel();
        });

        document.getElementById('exportPdfBtn').addEventListener('click', () => {
            this.exportToPDF();
        });

        // Report grouping
        document.getElementById('productionGroupBy').addEventListener('change', () => {
            this.updateProductionReport();
        });

        document.getElementById('packingGroupBy').addEventListener('change', () => {
            this.updatePackingReport();
        });
    }

    /**
     * יצירת קובץ דוגמה
     */
    createSampleFile() {
        if (typeof createSampleExcelFile === 'function') {
            createSampleExcelFile();
            alert('קובץ הדוגמה נוצר בהצלחה! ניתן לטעון אותו למערכת.');
        } else {
            alert('שגיאה ביצירת קובץ הדוגמה. אנא ודא שכל הקבצים נטענו.');
        }
    }

    /**
     * טעינת קובץ Excel
     */
    async handleFileUpload(file) {
        if (!file) return;

        this.showLoading(true);

        try {
            const data = await this.parseExcelFile(file);
            this.data = data;
            this.filteredData = [...data];
            
            this.setupFilters();
            this.renderResults();
            this.updateReports();
            
            this.showSections();
            
        } catch (error) {
            console.error('שגיאה בטעינת הקובץ:', error);
            alert('שגיאה בטעינת הקובץ. אנא ודא שהקובץ תקין.');
        } finally {
            this.showLoading(false);
        }
    }

    /**
     * פרסור קובץ Excel
     */
    async parseExcelFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    
                    const sheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[sheetName];
                    
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                    
                    if (jsonData.length < 2) {
                        throw new Error('הקובץ ריק או לא מכיל נתונים');
                    }

                    const headers = jsonData[0];
                    const rows = jsonData.slice(1);
                    
                    const processedData = rows.map(row => {
                        const obj = {};
                        headers.forEach((header, index) => {
                            if (header && header.trim()) {
                                obj[header.trim()] = row[index] || '';
                            }
                        });
                        
                        // חישוב שדות נוספים
                        this.computeRowCalculations(obj);
                        return obj;
                    });

                    resolve(processedData);
                    
                } catch (error) {
                    reject(error);
                }
            };
            
            reader.onerror = () => reject(new Error('שגיאה בקריאת הקובץ'));
            reader.readAsArrayBuffer(file);
        });
    }

    /**
     * חישוב שדות נוספים לכל שורה
     */
    computeRowCalculations(row) {
        // חישוב מנות לתכנון
        const quantity = parseFloat(row['כמות מוצר']) || 0;
        const param8 = parseFloat(row['פרמטר 8 למוצר']) || 0;
        
        let targetMeals = 0;
        let optimizationResult = null;

        // בדיקה אם צריך לבצע אופטימיזציה
        if (param8 > 0 && row['שיטת אירוז'] !== 'חמגשיות') {
            targetMeals = quantity / param8;
            
            // עיגול לפי כללים - תמיד עיגול למטה קודם
            let roundedTarget = Math.floor(targetMeals);
            
            console.log(`חישוב עבור ${row['מק"ט']}: כמות=${quantity}, פרמטר8=${param8}, מנות לתכנון=${targetMeals}, מעוגל=${roundedTarget}`);
            
            // נסה אופטימיזציה עם הערך המעוגל
            if (roundedTarget > 0) {
                optimizationResult = this.optimizePacks(roundedTarget, this.allowOverage);
                console.log(`תוצאת אופטימיזציה:`, optimizationResult);
            }
            
            // אם אין פתרון ויש עודף מותר, נסה עם עיגול למעלה
            if (!optimizationResult && this.allowOverage && targetMeals > 0) {
                roundedTarget = Math.ceil(targetMeals);
                optimizationResult = this.optimizePacks(roundedTarget, true);
                console.log(`תוצאת אופטימיזציה עם עודף:`, optimizationResult);
            }
        }

        // שמירת התוצאות
        row['מנות לתכנון'] = targetMeals;
        row['מנות לתכנון מעוגל'] = Math.floor(targetMeals);
        
        if (optimizationResult) {
            row['אריזות 5'] = optimizationResult.packs5;
            row['אריזות 7'] = optimizationResult.packs7;
            row['סה"כ אריזות'] = optimizationResult.packs5 + optimizationResult.packs7;
            row['עודף/פחת'] = optimizationResult.overage || 0;
            row['סטטוס אופטימיזציה'] = optimizationResult.exact ? 'מדויק' : 
                                         optimizationResult.overage > 0 ? 'עודף' : 'אין פתרון';
        } else {
            if (row['שיטת אירוז'] === 'חמגשיות') {
                row['סטטוס אופטימיזציה'] = 'דילג (חמגשיות)';
            } else if (param8 <= 0) {
                row['סטטוס אופטימיזציה'] = 'ללא פרמטר 8';
            } else {
                row['סטטוס אופטימיזציה'] = 'לא ניתן';
            }
            row['אריזות 5'] = 0;
            row['אריזות 7'] = 0;
            row['סה"כ אריזות'] = 0;
            row['עודף/פחת'] = 0;
        }

        // חישוב מיכלים וחמגשיות
        row['כמות מיכלים מחושב'] = parseFloat(row['כמות מיכלים']) || 0;
        row['סוג מיכל'] = row['קוד מיכל'] || '';
        row['שיטת אירוז מחושב'] = row['שיטת אירוז'] || '';
    }

    /**
     * אלגוריתם אופטימיזציית מארזים
     * @param {number} target - כמות מנות לתכנון
     * @param {boolean} allowOverage - האם לאפשר עודף
     * @returns {Object|null}
     */
    optimizePacks(target, allowOverage = false) {
        console.log(`אופטימיזציה עבור target=${target}, allowOverage=${allowOverage}`);
        
        // פונקציה לבדיקה אם ניתן לייצג מספר באמצעות 5 ו-7
        const canRepresent = (n) => {
            for (let a = 0; a <= Math.floor(n/5); a++) {
                const r = n - 5*a;
                if (r % 7 === 0) {
                    return {
                        packs5: a, 
                        packs7: r/7, 
                        exact: true, 
                        overage: 0
                    };
                }
            }
            return null;
        };

        // שלב 1: נסה התאמה מדויקת
        const exact = canRepresent(target);
        console.log(`התאמה מדויקת עבור ${target}:`, exact);
        if (exact) return exact;

        if (!allowOverage) {
            console.log(`אין עודף מותר, מחזיר null`);
            return null;
        }

        // שלב 2: אפשר מינימום פחת (עודף)
        let n = target + 1;
        const maxTry = target + 200; // גבול סביר
        
        while (n <= maxTry) {
            const sol = canRepresent(n);
            if (sol) {
                console.log(`מצאתי פתרון עם עודף עבור ${n}:`, sol);
                return {
                    ...sol, 
                    exact: false, 
                    overage: n - target
                };
            }
            n++;
        }
        
        console.log(`לא מצאתי פתרון עד ${maxTry}`);
        return null;
    }

    /**
     * חיפוש נתונים
     */
    handleSearch(query) {
        if (!query.trim()) {
            this.filteredData = [...this.data];
        } else {
            const searchTerm = query.toLowerCase();
            this.filteredData = this.data.filter(row => {
                // חיפוש לפי מספר הזמנה (שווה)
                const orderNumber = String(row['הזמנה'] || '').toLowerCase();
                if (orderNumber === searchTerm) {
                    return true;
                }

                // חיפוש לפי לקוח (מכיל)
                const customerNumber = String(row['מס. לקוח'] || '').toLowerCase();
                const customerName = String(row['שם לקוח'] || '').toLowerCase();
                const phone = String(row['מספר טלפון'] || '').toLowerCase();
                
                return customerNumber.includes(searchTerm) || 
                       customerName.includes(searchTerm) || 
                       phone.includes(searchTerm);
            });
        }

        this.applyFilters();
        this.renderResults();
        
        // הצגת כרטיס הזמנה אם חיפוש לפי מספר הזמנה
        if (query.trim() && this.filteredData.length > 0) {
            const orderNumber = String(this.filteredData[0]['הזמנה'] || '');
            if (orderNumber === query.trim()) {
                this.showOrderCard(this.filteredData[0]);
            }
        } else {
            this.hideOrderCard();
        }
    }

    /**
     * הצגת כרטיס הזמנה
     */
    showOrderCard(order) {
        const orderCard = document.getElementById('orderCard');
        const orderDetails = document.getElementById('orderDetails');
        
        orderDetails.innerHTML = `
            <div class="bg-blue-50 p-4 rounded-lg">
                <h3 class="font-bold text-lg mb-2">מספר הזמנה: ${order['הזמנה'] || ''}</h3>
                <p><strong>תאריך:</strong> ${order['תאריך'] || ''}</p>
                <p><strong>לקוח:</strong> ${order['שם לקוח'] || ''}</p>
            </div>
            <div class="bg-green-50 p-4 rounded-lg">
                <h4 class="font-bold mb-2">פרטי משלוח</h4>
                <p><strong>עיר:</strong> ${order['עיר'] || ''}</p>
                <p><strong>סניף:</strong> ${order['סניף'] || ''}</p>
                <p><strong>סוג:</strong> ${order['סוג'] || ''}</p>
            </div>
            <div class="bg-yellow-50 p-4 rounded-lg">
                <h4 class="font-bold mb-2">כשרות ופרטים</h4>
                <p><strong>כשרות:</strong> ${order['פרמטר 2 ללקוח'] || ''}</p>
                <p><strong>סוג לקוח:</strong> ${order['פרמטר 1 ללקוח'] || ''}</p>
            </div>
        `;
        
        orderCard.classList.remove('hidden');
    }

    /**
     * הסתרת כרטיס הזמנה
     */
    hideOrderCard() {
        document.getElementById('orderCard').classList.add('hidden');
    }

    /**
     * הגדרת סינונים
     */
    setupFilters() {
        const filterOptions = {
            branchFilter: [...new Set(this.data.map(row => row['סניף']).filter(Boolean))].sort(),
            cityFilter: [...new Set(this.data.map(row => row['עיר']).filter(Boolean))].sort(),
            customerTypeFilter: [...new Set(this.data.map(row => row['פרמטר 1 ללקוח']).filter(Boolean))].sort(),
            categoryFilter: [...new Set(this.data.map(row => row['פרמטר 1 לקוד']).filter(Boolean))].sort()
        };

        Object.keys(filterOptions).forEach(filterId => {
            const select = document.getElementById(filterId);
            // שמירת האופציה הראשונה (כל האפשרויות)
            const firstOption = select.querySelector('option[value=""]');
            select.innerHTML = '';
            select.appendChild(firstOption);
            
            filterOptions[filterId].forEach(option => {
                const optionElement = document.createElement('option');
                optionElement.value = option;
                optionElement.textContent = option;
                select.appendChild(optionElement);
            });
        });
    }

    /**
     * הפעלת סינונים
     */
    applyFilters() {
        const activeFilters = {
            branch: document.getElementById('branchFilter').value,
            city: document.getElementById('cityFilter').value,
            customerType: document.getElementById('customerTypeFilter').value,
            category: document.getElementById('categoryFilter').value
        };

        this.filteredData = this.data.filter(row => {
            return (!activeFilters.branch || row['סניף'] === activeFilters.branch) &&
                   (!activeFilters.city || row['עיר'] === activeFilters.city) &&
                   (!activeFilters.customerType || row['פרמטר 1 ללקוח'] === activeFilters.customerType) &&
                   (!activeFilters.category || row['פרמטר 1 לקוד'] === activeFilters.category);
        });

        this.renderResults();
        this.updateReports();
    }

    /**
     * ניקוי סינונים
     */
    clearFilters() {
        ['branchFilter', 'cityFilter', 'customerTypeFilter', 'categoryFilter'].forEach(id => {
            document.getElementById(id).value = '';
        });
        
        this.filteredData = [...this.data];
        this.renderResults();
        this.updateReports();
    }

    /**
     * הצגת תוצאות בטבלה
     */
    renderResults() {
        const tableBody = document.getElementById('tableBody');
        const tableHeader = document.getElementById('tableHeader');
        
        if (this.filteredData.length === 0) {
            tableBody.innerHTML = '<tr><td colspan="100%" class="text-center py-8 text-gray-500">אין נתונים להצגה</td></tr>';
            return;
        }

        // יצירת כותרות טבלה
        const headers = [
            'הזמנה', 'תאריך', 'מס. לקוח', 'שם לקוח', 'פרמטר 1 ללקוח', 'פרמטר 2 ללקוח',
            'עיר', 'סניף', 'מק"ט', 'תאור מוצר', 'פרמטר 1 לקוד', 'פרמטר 6 למוצר',
            'כמות מוצר', 'פרמטר 8 למוצר', 'מנות לתכנון', 'אריזות 5', 'אריזות 7',
            'סה"כ אריזות', 'עודף/פחת', 'סטטוס אופטימיזציה', 'שיטת אירוז', 'כמות מיכלים'
        ];

        tableHeader.innerHTML = headers.map(header => 
            `<th class="px-4 py-2 text-right border-b border-gray-200">${header}</th>`
        ).join('');

        // יצירת שורות טבלה
        tableBody.innerHTML = this.filteredData.map(row => {
            return `<tr class="border-b border-gray-100 hover:bg-gray-50">
                ${headers.map(header => {
                    let value = row[header] || '';
                    
                    // עיגול מספרים מחושבים
                    if (['כמות מוצר', 'מנות לתכנון', 'אריזות 5', 'אריזות 7', 'סה"כ אריזות', 'עודף/פחת', 'כמות מיכלים'].includes(header)) {
                        value = Math.round(parseFloat(value) || 0);
                    }
                    
                    return `<td class="px-4 py-2 text-right">${value}</td>`;
                }).join('')}
            </tr>`;
        }).join('');
    }

    /**
     * עדכון דוחות
     */
    updateReports() {
        this.updateProductionReport();
        this.updatePackingReport();
    }

    /**
     * עדכון דוח ייצור
     */
    updateProductionReport() {
        const groupBy = document.getElementById('productionGroupBy').value;
        const grouped = this.groupData(groupBy);
        
        const tableBody = document.getElementById('productionTableBody');
        tableBody.innerHTML = Object.entries(grouped).map(([key, rows]) => {
            // סיכום כמות מוצר
            const totalQuantity = rows.reduce((sum, row) => sum + (parseFloat(row['כמות מוצר']) || 0), 0);
            
            // סיכום כמות מנות לפריט (מספר מנות לשורה)
            const totalMealsPerItem = rows.reduce((sum, row) => sum + (parseFloat(row['מספר מנות לשורה']) || 0), 0);
            
            // סיכום סה"כ מנות היום (מספר מנות כללי) - רק מהפריט הראשון בכל הזמנה
            const orderGroups = {};
            rows.forEach(row => {
                const orderNum = row['הזמנה'];
                if (!orderGroups[orderNum]) {
                    orderGroups[orderNum] = [];
                }
                orderGroups[orderNum].push(row);
            });
            
            const totalMealsToday = Object.values(orderGroups).reduce((sum, orderRows) => {
                // לוקחים רק את הפריט הראשון בכל הזמנה
                const firstRow = orderRows[0];
                return sum + (parseFloat(firstRow['מספר מנות כללי']) || 0);
            }, 0);
            
            // סיכום מנות אלרגניות - רק מהפריט הראשון בכל הזמנה
            const totalAllergenic = Object.values(orderGroups).reduce((sum, orderRows) => {
                const firstRow = orderRows[0];
                return sum + (parseFloat(firstRow['מספר מנות אלרגניות']) || 0);
            }, 0);
            
            // סיכום מנות צמחוניות - רק מהפריט הראשון בכל הזמנה
            const totalVegetarian = Object.values(orderGroups).reduce((sum, orderRows) => {
                const firstRow = orderRows[0];
                return sum + (parseFloat(firstRow['מספר מנות צמחוניות']) || 0);
            }, 0);
            
            return `<tr class="border-b border-gray-100">
                <td class="px-4 py-2 text-right font-medium">${key}</td>
                <td class="px-4 py-2 text-right">${Math.round(totalQuantity)}</td>
                <td class="px-4 py-2 text-right">${Math.round(totalMealsPerItem)}</td>
                <td class="px-4 py-2 text-right">${Math.round(totalMealsToday)}</td>
                <td class="px-4 py-2 text-right">${Math.round(totalAllergenic)}</td>
                <td class="px-4 py-2 text-right">${Math.round(totalVegetarian)}</td>
            </tr>`;
        }).join('');
    }

    /**
     * עדכון דוח אריזה
     */
    updatePackingReport() {
        const groupBy = document.getElementById('packingGroupBy').value;
        const grouped = this.groupData(groupBy);
        
        const tableBody = document.getElementById('packingTableBody');
        tableBody.innerHTML = Object.entries(grouped).map(([key, rows]) => {
            const totalPacks5 = rows.reduce((sum, row) => sum + (parseFloat(row['אריזות 5']) || 0), 0);
            const totalPacks7 = rows.reduce((sum, row) => sum + (parseFloat(row['אריזות 7']) || 0), 0);
            const totalPacks = totalPacks5 + totalPacks7;
            
            // סיכום מיכלים (לא חמגשיות)
            const totalContainers = rows
                .filter(row => row['שיטת אירוז'] !== 'חמגשיות')
                .reduce((sum, row) => sum + (parseFloat(row['כמות מיכלים']) || 0), 0);
            
            // סיכום חמגשיות
            const totalTrays = rows
                .filter(row => row['שיטת אירוז'] === 'חמגשיות')
                .reduce((sum, row) => sum + (parseFloat(row['כמות מיכלים']) || 0), 0);
            
            const noSolutionCount = rows.filter(row => row['סטטוס אופטימיזציה'] === 'אין פתרון').length;
            const totalOverage = rows.reduce((sum, row) => sum + (parseFloat(row['עודף/פחת']) || 0), 0);
            
            return `<tr class="border-b border-gray-100">
                <td class="px-4 py-2 text-right font-medium">${key}</td>
                <td class="px-4 py-2 text-right">${Math.round(totalPacks5)}</td>
                <td class="px-4 py-2 text-right">${Math.round(totalPacks7)}</td>
                <td class="px-4 py-2 text-right">${Math.round(totalPacks)}</td>
                <td class="px-4 py-2 text-right">${Math.round(totalContainers)}</td>
                <td class="px-4 py-2 text-right">${Math.round(totalTrays)}</td>
                <td class="px-4 py-2 text-right">${Math.round(noSolutionCount)}</td>
                <td class="px-4 py-2 text-right">${Math.round(totalOverage)}</td>
            </tr>`;
        }).join('');
    }

    /**
     * קיבוץ נתונים לפי שדה
     */
    groupData(groupBy) {
        const grouped = {};
        
        this.filteredData.forEach(row => {
            const key = row[this.getGroupByField(groupBy)] || 'ללא ערך';
            if (!grouped[key]) {
                grouped[key] = [];
            }
            grouped[key].push(row);
        });
        
        return grouped;
    }

    /**
     * קבלת שדה קיבוץ
     */
    getGroupByField(groupBy) {
        const mapping = {
            'sku': 'מק"ט',
            'product': 'תאור מוצר',
            'category': 'פרמטר 1 לקוד',
            'department': 'פרמטר 6 למוצר',
            'branch': 'סניף'
        };
        return mapping[groupBy] || 'מק"ט';
    }

    /**
     * חישוב מחדש של אופטימיזציות
     */
    recalculateOptimizations() {
        this.data.forEach(row => {
            this.computeRowCalculations(row);
        });
        
        this.renderResults();
        this.updateReports();
    }

    /**
     * ייצוא ל-Excel
     */
    exportToExcel() {
        if (this.filteredData.length === 0) {
            alert('אין נתונים לייצוא');
            return;
        }

        const worksheet = XLSX.utils.json_to_sheet(this.filteredData);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'הזמנות');
        
        XLSX.writeFile(workbook, `הזמנות_${new Date().toISOString().split('T')[0]}.xlsx`);
    }

    /**
     * ייצוא ל-PDF
     */
    exportToPDF() {
        if (this.filteredData.length === 0) {
            alert('אין נתונים לייצוא');
            return;
        }

        const { jsPDF } = window.jspdf;
        const doc = new jsPDF('l', 'mm', 'a4');
        
        // כותרת
        doc.setFontSize(16);
        doc.text('דוח הזמנות', 20, 20);
        
        // טבלת נתונים
        const columns = [
            'הזמנה', 'תאריך', 'לקוח', 'מוצר', 'כמות', 'אריזות 5', 'אריזות 7', 'סטטוס'
        ];
        
        const rows = this.filteredData.map(row => [
            row['הזמנה'] || '',
            row['תאריך'] || '',
            row['שם לקוח'] || '',
            row['תאור מוצר'] || '',
            row['כמות מוצר'] || '',
            row['אריזות 5'] || '',
            row['אריזות 7'] || '',
            row['סטטוס אופטימיזציה'] || ''
        ]);
        
        doc.autoTable({
            columns: columns,
            body: rows,
            startY: 30,
            styles: { fontSize: 8, halign: 'right' },
            headStyles: { fillColor: [66, 139, 202] }
        });
        
        doc.save(`הזמנות_${new Date().toISOString().split('T')[0]}.pdf`);
    }

    /**
     * הצגת/הסתרת סקציות
     */
    showSections() {
        document.getElementById('filtersSection').classList.remove('hidden');
        document.getElementById('resultsSection').classList.remove('hidden');
        document.getElementById('reportsSection').classList.remove('hidden');
    }

    /**
     * הצגת/הסתרת מסך טעינה
     */
    showLoading(show) {
        const overlay = document.getElementById('loadingOverlay');
        if (show) {
            overlay.classList.remove('hidden');
        } else {
            overlay.classList.add('hidden');
        }
    }

    /**
     * שמירת הגדרות
     */
    saveSettings() {
        const settings = {
            allowOverage: this.allowOverage,
            filters: this.filters
        };
        localStorage.setItem('orderManagementSettings', JSON.stringify(settings));
    }

    /**
     * טעינת הגדרות
     */
    loadSettings() {
        const saved = localStorage.getItem('orderManagementSettings');
        if (saved) {
            try {
                const settings = JSON.parse(saved);
                this.allowOverage = settings.allowOverage || false;
                this.filters = settings.filters || {};
                
                // עדכון UI
                document.getElementById('allowOverageToggle').checked = this.allowOverage;
            } catch (error) {
                console.error('שגיאה בטעינת הגדרות:', error);
            }
        }
    }
}

// אתחול המערכת
document.addEventListener('DOMContentLoaded', () => {
    new OrderManagementSystem();
});
