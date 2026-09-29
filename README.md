<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>記帳與預算管理系統</title>
    <!-- 引入 Tailwind CSS -->
    <script src="https://cdn.jsdelivr.net/npm/@tailwindcss/browser@4"></script>
    <!-- 引入 SheetJS (用於匯出 Excel 檔案) -->
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <!-- 引入 html-docx-js 與 FileSaver (用於匯出 Word 檔案) -->
    <script src="https://cdn.jsdelivr.net/npm/html-docx-js@0.3.1/dist/html-docx.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/file-saver@2.0.5/dist/FileSaver.min.js"></script>
</head>
<body class="bg-gray-100 text-gray-800 font-sans p-4">

    <div class="max-w-md mx-auto bg-white rounded-lg shadow-md overflow-hidden pb-6">
        <!-- 標題與備份功能列 -->
        <header class="bg-indigo-600 text-white p-4 flex flex-col gap-2">
            <div class="flex justify-between items-center">
                <h1 class="font-bold text-lg">記帳與預算管理系統</h1>
                <div class="space-x-1">
                    <button onclick="exportData()" class="bg-indigo-700 hover:bg-indigo-800 text-xs px-2 py-1 rounded transition" title="下載資料備份 JSON 檔">備份</button>
                    <label class="bg-indigo-700 hover:bg-indigo-800 text-xs px-2 py-1 rounded transition cursor-pointer inline-block">
                        還原 <input type="file" id="importFile" accept=".json" onchange="importData(event)" class="hidden">
                    </label>
                </div>
            </div>
            <!-- 報表匯出按鈕列 -->
            <div class="flex gap-2 pt-1 border-t border-indigo-500 text-xs">
                <button onclick="exportToExcel()" class="flex-1 bg-emerald-600 hover:bg-emerald-700 py-1.5 rounded font-semibold transition text-center">
                    📊 匯出 Excel
                </button>
                <button onclick="exportToWord()" class="flex-1 bg-blue-600 hover:bg-blue-700 py-1.5 rounded font-semibold transition text-center">
                    📄 匯出 Word
                </button>
            </div>
        </header>

        <!-- 1. 預算與儲值管理區 -->
        <div class="p-4 bg-gray-50 border-b space-y-3">
            <h2 class="text-xs font-bold text-gray-700">各成員預算與儲值管理</h2>
            
            <!-- Sasa 區塊 -->
            <div class="bg-white p-3 border rounded shadow-sm space-y-2">
                <div class="flex justify-between items-center text-xs">
                    <span class="font-bold text-indigo-900">Sasa 總預算：<span id="displayBudgetSasa" class="text-green-600 text-sm font-bold">$1000</span></span>
                    <button onclick="editBaseBudget('Sasa')" class="text-gray-400 hover:text-indigo-600">修改初始</button>
                </div>
                <div class="flex gap-2">
                    <input type="number" id="topupSasa" placeholder="追加儲值金額" class="w-full p-1 border rounded text-xs bg-white">
                    <button onclick="topUpBudget('Sasa')" class="bg-green-600 text-white px-3 py-1 rounded text-xs font-semibold hover:bg-green-700 whitespace-nowrap">+ 儲值</button>
                </div>
            </div>

            <!-- 毓宸 區塊 -->
            <div class="bg-white p-3 border rounded shadow-sm space-y-2">
                <div class="flex justify-between items-center text-xs">
                    <span class="font-bold text-indigo-900">毓宸 總預算：<span id="displayBudgetYuchen" class="text-green-600 text-sm font-bold">$1000</span></span>
                    <button onclick="editBaseBudget('毓宸')" class="text-gray-400 hover:text-indigo-600">修改初始</button>
                </div>
                <div class="flex gap-2">
                    <input type="number" id="topupYuchen" placeholder="追加儲值金額" class="w-full p-1 border rounded text-xs bg-white">
                    <button onclick="topUpBudget('毓宸')" class="bg-green-600 text-white px-3 py-1 rounded text-xs font-semibold hover:bg-green-700 whitespace-nowrap">+ 儲值</button>
                </div>
            </div>
        </div>

        <!-- 2. 新增消費表單 -->
        <div class="p-4 bg-indigo-50 border-b">
            <h2 class="text-xs font-bold text-indigo-900 mb-2">新增一筆消費紀錄</h2>
            <form id="expenseForm" class="space-y-3">
                <div>
                    <label class="block text-xs font-semibold text-gray-600 mb-1">購買日期</label>
                    <input type="date" id="date" required class="w-full p-2 border rounded text-sm bg-white">
                </div>

                <div>
                    <label class="block text-xs font-semibold text-gray-600 mb-1">購買人</label>
                    <div class="space-y-2">
                        <select id="buyerSelect" onchange="handleBuyerChange()" class="w-full p-2 border rounded text-sm bg-white">
                            <option value="Sasa">Sasa</option>
                            <option value="毓宸">毓宸</option>
                            <option value="custom">-- 自行輸入其他人名 --</option>
                        </select>
                        <input type="text" id="buyerCustom" placeholder="請輸入名字" class="w-full p-2 border rounded text-sm bg-white hidden">
                    </div>
                </div>

                <div>
                    <label class="block text-xs font-semibold text-gray-600 mb-1">購買物品</label>
                    <input type="text" id="item" placeholder="例: 便當 (自助餐、麵線...)" required class="w-full p-2 border rounded text-sm bg-white">
                </div>

                <div>
                    <label class="block text-xs font-semibold text-gray-600 mb-1">金額 (元)</label>
                    <input type="number" id="price" placeholder="59" required class="w-full p-2 border rounded text-sm bg-white">
                </div>

                <button type="submit" class="w-full bg-indigo-600 text-white py-2 rounded text-sm font-semibold hover:bg-indigo-700 transition">
                    新增紀錄
                </button>
            </form>
        </div>

        <!-- 3. 列表與剩餘顯示區 -->
        <div class="p-4">
            <div class="flex justify-between items-center mb-3">
                <h2 class="text-sm font-bold text-gray-700">剩餘額度總覽與明細</h2>
                <button onclick="resetData()" class="text-xs text-red-500 hover:underline">清空全部資料</button>
            </div>

            <!-- 剩餘總覽卡片 -->
            <div id="summaryCard" class="grid grid-cols-2 gap-2 mb-4">
                <!-- 動態載入 -->
            </div>

            <!-- 歷史清單 -->
            <div id="expenseList" class="space-y-2">
                <!-- 動態載入 -->
            </div>
        </div>
    </div>

    <script>
        // 設定預設日期為今天
        document.getElementById('date').valueAsDate = new Date();

        // 安全讀取 localStorage
        let budgets = JSON.parse(localStorage.getItem('app_budgets')) || { "Sasa": 1000, "毓宸": 1000 };
        let expenses = JSON.parse(localStorage.getItem('app_expenses')) || [];

        const expenseForm = document.getElementById('expenseForm');
        const expenseList = document.getElementById('expenseList');
        const summaryCard = document.getElementById('summaryCard');
        const buyerSelect = document.getElementById('buyerSelect');
        const buyerCustom = document.getElementById('buyerCustom');

        function handleBuyerChange() {
            if (buyerSelect.value === 'custom') {
                buyerCustom.classList.remove('hidden');
                buyerCustom.focus();
            } else {
                buyerCustom.classList.add('hidden');
                buyerCustom.value = '';
            }
        }

        function editBaseBudget(person) {
            let current = budgets[person] || 0;
            let val = prompt(`請輸入 ${person} 的初始預算金額：`, current);
            if (val !== null) {
                let num = parseFloat(val);
                if (!isNaN(num)) {
                    budgets[person] = num;
                    localStorage.setItem('app_budgets', JSON.stringify(budgets));
                    renderUI();
                } else {
                    alert('請輸入有效的數字！');
                }
            }
        }

        function topUpBudget(person) {
            let inputId = person === 'Sasa' ? 'topupSasa' : 'topupYuchen';
            let val = parseFloat(document.getElementById(inputId).value);
            if (isNaN(val) || val <= 0) {
                alert('請輸入有效的儲值金額！');
                return;
            }
            budgets[person] = (budgets[person] || 0) + val;
            localStorage.setItem('app_budgets', JSON.stringify(budgets));
            document.getElementById(inputId).value = '';
            renderUI();
            alert(`成功幫 ${person} 儲值 $${val}！`);
        }

        // 匯出資料備份 (JSON)
        function exportData() {
            const data = { budgets: budgets, expenses: expenses };
            const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(data, null, 2));
            const downloadAnchor = document.createElement('a');
            downloadAnchor.setAttribute("href", dataStr);
            downloadAnchor.setAttribute("download", "expense_backup_" + new Date().toISOString().slice(0,10) + ".json");
            document.body.appendChild(downloadAnchor);
            downloadAnchor.click();
            downloadAnchor.remove();
        }

        // 匯入還原備份 (JSON)
        function importData(event) {
            const file = event.target.files[0];
            if (!file) return;
            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const imported = JSON.parse(e.target.result);
                    if (imported.budgets && imported.expenses) {
                        budgets = imported.budgets;
                        expenses = imported.expenses;
                        localStorage.setItem('app_budgets', JSON.stringify(budgets));
                        localStorage.setItem('app_expenses', JSON.stringify(expenses));
                        renderUI();
                        alert('資料還原成功！');
                    } else {
                        alert('備份檔案格式不正確！');
                    }
                } catch (err) {
                    alert('解析備份檔案失敗！');
                }
            };
            reader.readAsText(file);
        }

        // 取得所有參與人清單
        function getAllPeople() {
            let allPeopleSet = new Set(["Sasa", "毓宸", ...Object.keys(budgets)]);
            expenses.forEach(item => {
                if(item.buyer) allPeopleSet.add(item.buyer);
            });
            return Array.from(allPeopleSet);
        }

        // 📊 匯出 Excel 檔案
        function exportToExcel() {
            const wb = XLSX.utils.book_new();

            // 1. 總覽資料
            let summaryData = [["購買人", "總預算", "已花費金額", "剩餘額度"]];
            getAllPeople().forEach(person => {
                let spent = expenses.filter(i => i.buyer === person).reduce((s, i) => s + i.price, 0);
                let initial = budgets[person] || 0;
                let remaining = initial - spent;
                summaryData.push([person, initial, spent, remaining]);
            });
            const wsSummary = XLSX.utils.aoa_to_sheet(summaryData);
            XLSX.utils.book_append_sheet(wb, wsSummary, "預算總覽");

            // 2. 消費明細資料
            let detailData = [["購買日期", "購買人", "購買物品", "金額"]];
            expenses.forEach(i => {
                detailData.push([i.date, i.buyer, i.item, i.price]);
            });
            const wsDetail = XLSX.utils.aoa_to_sheet(detailData);
            XLSX.utils.book_append_sheet(wb, wsDetail, "消費明細");

            // 下載檔案
            XLSX.writeFile(wb, "記帳與預算報表_" + new Date().toISOString().slice(0,10) + ".xlsx");
        }

        // 📄 匯出 Word 檔案
        function exportToWord() {
            let people = getAllPeople();
            
            let htmlContent = `
                <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
                <head><meta charset='utf-8'><title>記帳報表</title></head>
                <body style="font-family: 'Microsoft JhengHei', sans-serif;">
                    <h1 style="color: #4F46E5; text-align: center;">記帳與預算管理系統報表</h1>
                    <p style="text-align: center; color: #666; font-size: 12px;">匯出日期：${new Date().toLocaleDateString()}</p>
                    
                    <h2 style="color: #333; border-bottom: 2px solid #4F46E5; padding-bottom: 4px;">一、 預算與剩餘額度總覽</h2>
                    <table border="1" cellspacing="0" cellpadding="6" style="width: 100%; border-collapse: collapse; text-align: center; font-size: 14px;">
                        <tr style="background-color: #4F46E5; color: white;">
                            <th>購買人</th><th>總預算</th><th>已花費金額</th><th>剩餘額度</th>
                        </tr>
            `;

            people.forEach(person => {
                let spent = expenses.filter(i => i.buyer === person).reduce((s, i) => s + i.price, 0);
                let initial = budgets[person] || 0;
                let remaining = initial - spent;
                htmlContent += `
                    <tr>
                        <td>${person}</td>
                        <td>$${initial}</td>
                        <td>$${spent}</td>
                        <td style="color: ${remaining >= 0 ? 'green' : 'red'}; font-weight: bold;">$${remaining}</td>
                    </tr>
                `;
            });

            htmlContent += `
                    </table>
                    <br><br>
                    <h2 style="color: #333; border-bottom: 2px solid #4F46E5; padding-bottom: 4px;">二、 消費明細紀錄</h2>
                    <table border="1" cellspacing="0" cellpadding="6" style="width: 100%; border-collapse: collapse; text-align: center; font-size: 14px;">
                        <tr style="background-color: #4F46E5; color: white;">
                            <th>購買日期</th><th>購買人</th><th>購買物品</th><th>金額</th>
                        </tr>
            `;

            if (expenses.length === 0) {
                htmlContent += `<tr><td colspan="4" style="color: #888;">目前沒有任何消費紀錄</td></tr>`;
            } else {
                expenses.slice().reverse().forEach(item => {
                    htmlContent += `
                        <tr>
                            <td>${item.date}</td>
                            <td>${item.buyer}</td>
                            <td style="text-align: left;">${item.item}</td>
                            <td style="text-align: right;">$${item.price}</td>
                        </tr>
                    `;
                });
            }

            htmlContent += `
                    </table>
                </body>
                </html>
            `;

            let converted = htmlDocx.asBlob(htmlContent);
            saveAs(converted, "記帳與預算報表_" + new Date().toISOString().slice(0,10) + ".docx");
        }

        // 渲染畫面主程式
        function renderUI() {
            try {
                document.getElementById('displayBudgetSasa').textContent = `$${budgets["Sasa"] || 0}`;
                document.getElementById('displayBudgetYuchen').textContent = `$${budgets["毓宸"] || 0}`;

                let allPeopleSet = new Set(["Sasa", "毓宸", ...Object.keys(budgets)]);
                expenses.forEach(item => {
                    if(item.buyer) allPeopleSet.add(item.buyer);
                });

                // 1. 顯示各自剩餘多少
                summaryCard.innerHTML = '';
                allPeopleSet.forEach(person => {
                    let spent = expenses
                        .filter(item => item.buyer === person)
                        .reduce((sum, item) => sum + item.price, 0);
                    
                    let initial = budgets[person] || 0;
                    let remaining = initial - spent;

                    let card = document.createElement('div');
                    card.className = "bg-gray-50 p-3 border rounded text-center shadow-sm";
                    card.innerHTML = `
                        <div class="text-xs font-bold text-gray-600">${person} 剩餘</div>
                        <div class="text-lg font-bold ${remaining >= 0 ? 'text-green-600' : 'text-red-600'}">$${remaining}</div>
                        <div class="text-[10px] text-gray-400 mt-1">總預算: $${initial} / 已花: $${spent}</div>
                    `;
                    summaryCard.appendChild(card);
                });

                // 2. 渲染購買清單
                expenseList.innerHTML = '';
                if (!expenses || expenses.length === 0) {
                    expenseList.innerHTML = '<p class="text-gray-400 text-center py-4 text-xs">目前沒有任何消費紀錄</p>';
                    return;
                }

                expenses.slice().reverse().forEach((item, index) => {
                    let actualIndex = expenses.length - 1 - index;
                    let div = document.createElement('div');
                    div.className = "bg-white p-3 border rounded text-xs space-y-1 shadow-sm";
                    div.innerHTML = `
                        <div class="flex justify-between text-gray-400">
                            <span>📅 ${item.date}</span>
                            <button onclick="deleteItem(${actualIndex})" class="text-red-400 hover:text-red-600 font-bold">刪除</button>
                        </div>
                        <div class="font-bold text-gray-800">購買人：<span class="text-indigo-600">${item.buyer}</span></div>
                        <div>購買物品：${item.item}</div>
                        <div class="font-semibold text-gray-700">金額：$${item.price}</div>
                    `;
                    expenseList.appendChild(div);
                });
            } catch (err) {
                console.error("渲染錯誤:", err);
            }
        }

        // 新增紀錄事件
        expenseForm.addEventListener('submit', function(e) {
            e.preventDefault();
            
            let buyerName = buyerSelect.value;
            if (buyerName === 'custom') {
                buyerName = buyerCustom.value.trim();
            }

            const newItem = {
                date: document.getElementById('date').value,
                buyer: buyerName,
                item: document.getElementById('item').value.trim(),
                price: parseFloat(document.getElementById('price').value)
            };

            if (!newItem.item || isNaN(newItem.price) || !buyerName) {
                alert('請完整填寫購買人、物品與金額！');
                return;
            }

            if (budgets[buyerName] === undefined) {
                budgets[buyerName] = 0;
                localStorage.setItem('app_budgets', JSON.stringify(budgets));
            }

            expenses.push(newItem);
            localStorage.setItem('app_expenses', JSON.stringify(expenses));
            
            renderUI();
            expenseForm.reset();
            document.getElementById('date').valueAsDate = new Date();
            buyerSelect.value = 'Sasa';
            buyerCustom.classList.add('hidden');
            buyerCustom.value = '';
        });

        // 刪除單筆紀錄
        function deleteItem(index) {
            if(confirm('確定要刪除這筆紀錄嗎？')) {
                expenses.splice(index, 1);
                localStorage.setItem('app_expenses', JSON.stringify(expenses));
                renderUI();
            }
        }

        // 清空全部資料
        function resetData() {
            if(confirm('確定要清空全部的預算與消費紀錄嗎？')) {
                localStorage.removeItem('app_expenses');
                localStorage.removeItem('app_budgets');
                expenses = [];
                budgets = { "Sasa": 1000, "毓宸": 1000 };
                renderUI();
            }
        }

        // 初始化執行
        renderUI();
    </script>
</body>
</html>
