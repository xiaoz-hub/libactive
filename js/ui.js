// 分页相关变量
let currentPage = 1;
let pageSize = 10;
let totalPages = 1;

// 格式化问题描述，添加高亮颜色
function formatIssues(issues) {
    if (!issues || issues === '通过') return issues;
    
    // 将问题描述按分号分割
    const issueParts = issues.split(';');
    const formattedParts = issueParts.map(part => {
        const trimmedPart = part.trim();
        if (!trimmedPart) return '';
        
        // 为不同类型的问题添加不同的颜色
        if (trimmedPart.includes('学号格式不符合规则')) {
            return `<span class="text-red-600 font-semibold">${trimmedPart}</span>`;
        } else if (trimmedPart.includes('学号位数不符合')) {
            return `<span class="text-red-600 font-semibold">${trimmedPart}</span>`;
        } else if (trimmedPart.includes('未填写部门')) {
            return `<span class="text-orange-600 font-semibold">${trimmedPart}</span>`;
        } else if (trimmedPart.includes('部门名称未在规则列表中')) {
            return `<span class="text-purple-600 font-semibold">${trimmedPart}</span>`;
        } else if (trimmedPart.includes('活动名称未在规则列表中')) {
            return `<span class="text-blue-600 font-semibold">${trimmedPart}</span>`;
        } else if (trimmedPart.includes('活动分数不匹配')) {
            return `<span class="text-pink-600 font-semibold">${trimmedPart}</span>`;
        } else if (trimmedPart.includes('未填写活动分数')) {
            return `<span class="text-yellow-600 font-semibold">${trimmedPart}</span>`;
        } else if (trimmedPart.includes('缺少学号信息')) {
            return `<span class="text-red-600 font-semibold">${trimmedPart}</span>`;
        } else if (trimmedPart.includes('缺少姓名信息')) {
            return `<span class="text-red-600 font-semibold">${trimmedPart}</span>`;
        } else if (trimmedPart.includes('缺少活动名称信息')) {
            return `<span class="text-red-600 font-semibold">${trimmedPart}</span>`;
        } else if (trimmedPart.includes('活动规则未添加')) {
            return `<span class="text-red-600 font-semibold">${trimmedPart}</span>`;
        } else if (trimmedPart.includes('学号规则未添加')) {
            return `<span class="text-red-600 font-semibold">${trimmedPart}</span>`;
        } else if (trimmedPart.includes('部门规则未添加')) {
            return `<span class="text-red-600 font-semibold">${trimmedPart}</span>`;
        } else {
            return `<span class="text-gray-700">${trimmedPart}</span>`;
        }
    });
    
    return formattedParts.join('; ');
}

window.updateResultsTable = function() {
    resultsTable.innerHTML = '';
    let filteredResults = reviewResults;
    if (currentFilter === 'passed') filteredResults = reviewResults.filter(r => r.passed);
    else if (currentFilter === 'failed') filteredResults = reviewResults.filter(r => !r.passed);
    // 查询条件过滤
    if (typeof currentQuery !== 'undefined' && currentQuery) {
        if (currentQuery.department) {
            filteredResults = filteredResults.filter(r => (r.department || '').includes(currentQuery.department));
        }
        if (currentQuery.studentPrefix) {
            filteredResults = filteredResults.filter(r => (r.studentId || '').startsWith(currentQuery.studentPrefix));
        }
    }
    
    if (filteredResults.length === 0) {
        clearResultsBtn.classList.add('hidden');
        document.getElementById('paginationContainer').classList.add('hidden');
        const emptyRow = document.createElement('tr');
        emptyRow.innerHTML = `
            <td colspan="8" class="px-4 py-8 text-center text-gray-500">
                <i class="fa fa-info-circle text-2xl mb-2 block opacity-50"></i>
                ${currentFilter === 'passed' ? '没有通过的记录' : currentFilter === 'failed' ? '没有未通过的记录' : '没有记录'}
            </td>
        `;
        resultsTable.appendChild(emptyRow);
        return;
    }
    
    // 计算分页
    totalPages = Math.ceil(filteredResults.length / pageSize);
    if (currentPage > totalPages) currentPage = totalPages;
    if (currentPage < 1) currentPage = 1;
    
    // 如果当前页超出范围，重置到第一页
    if (currentPage > totalPages && totalPages > 0) {
        currentPage = 1;
    }
    
    const startIndex = (currentPage - 1) * pageSize;
    const endIndex = Math.min(startIndex + pageSize, filteredResults.length);
    const pageResults = filteredResults.slice(startIndex, endIndex);
    
    // 渲染当前页数据
    pageResults.forEach((result, index) => {
        const row = document.createElement('tr');
        row.className = index % 2 === 0 ? 'bg-white' : 'bg-gray-50';
        // 为审查结果添加圆角底色边框
        const resultBadge = result.passed 
            ? '<span class="inline-block px-3 py-1 text-xs font-medium bg-green-100 text-green-800 rounded-full border border-green-200">通过</span>'
            : '<span class="inline-block px-3 py-1 text-xs font-medium bg-red-100 text-red-800 rounded-full border border-red-200">未通过</span>';
        
        const deptDisplay = (result.department && result.department !== '未知') ? result.department : '<span class="text-red-600">未填写部门</span>';
        const sequenceNumber = startIndex + index + 1;
        
        // 格式化问题描述，添加高亮颜色
        const formattedIssues = formatIssues(result.issues);
        
        row.innerHTML = `
            <td class="px-4 py-2 text-sm text-gray-700 text-center">${sequenceNumber}</td>
            <td class="px-4 py-2 text-sm text-gray-700">${result.name}</td>
            <td class="px-4 py-2 text-sm text-gray-700">${result.studentId}</td>
            <td class="px-4 py-2 text-sm text-gray-700">${deptDisplay}</td>
            <td class="px-4 py-2 text-sm text-gray-700">${result.activityName}</td>
            <td class="px-4 py-2 text-sm text-gray-700">${result.score}</td>
            <td class="px-4 py-2 text-sm text-center">${resultBadge}</td>
            <td class="px-4 py-2 text-sm text-gray-700">${formattedIssues}</td>
        `;
        resultsTable.appendChild(row);
    });
    
    // 更新分页控件
    updatePagination(filteredResults.length);
    clearResultsBtn.classList.remove('hidden');
    
    // 更新统计信息（显示筛选后的结果）
    const passedCount = filteredResults.filter(r => r.passed).length;
    const failedCount = filteredResults.filter(r => !r.passed).length;
    totalRecords.textContent = filteredResults.length;
    passedRecords.textContent = passedCount;
    failedRecords.textContent = failedCount;
}

function updatePagination(totalRecords) {
    const paginationContainer = document.getElementById('paginationContainer');
    const pageSizeSelect = document.getElementById('pageSizeSelect');
    const currentPageInput = document.getElementById('currentPageInput');
    const totalPagesSpan = document.getElementById('totalPages');
    const startRecordSpan = document.getElementById('startRecord');
    const endRecordSpan = document.getElementById('endRecord');
    const totalRecordsDisplaySpan = document.getElementById('totalRecordsDisplay');
    
    // 显示分页控件
    paginationContainer.classList.remove('hidden');
    
    // 更新分页信息
    totalPagesSpan.textContent = totalPages;
    currentPageInput.value = currentPage;
    startRecordSpan.textContent = (currentPage - 1) * pageSize + 1;
    endRecordSpan.textContent = Math.min(currentPage * pageSize, totalRecords);
    totalRecordsDisplaySpan.textContent = totalRecords;
    
    // 更新按钮状态
    document.getElementById('firstPageBtn').disabled = currentPage === 1;
    document.getElementById('prevPageBtn').disabled = currentPage === 1;
    document.getElementById('nextPageBtn').disabled = currentPage === totalPages;
    document.getElementById('lastPageBtn').disabled = currentPage === totalPages;
    
    // 更新每页记录数选择器
    pageSizeSelect.value = pageSize;
}

window.goToPage = function(page) {
    if (page < 1 || page > totalPages) return;
    currentPage = page;
    updateResultsTable();
}

window.changePageSize = function(newSize) {
    pageSize = parseInt(newSize);
    currentPage = 1; // 重置到第一页
    updateResultsTable();
}

window.displayReviewResults = function() {
    totalRecords.textContent = reviewResults.length;
    const passedCount = reviewResults.filter(r => r.passed).length;
    passedRecords.textContent = passedCount;
    failedRecords.textContent = reviewResults.length - passedCount;
    
    // 重置分页状态
    currentPage = 1;
    
    updateResultsTable();
    document.getElementById('exportOptions').classList.remove('hidden');
    if (reviewResults.length === 0) {
        resultsTable.innerHTML = `
            <tr class="text-center">
                <td colspan="8" class="px-4 py-8 text-gray-500">
                <i class="fa fa-info-circle text-2xl mb-2 block opacity-50"></i>
                没有找到审查结果
            </td>
        `;
        document.getElementById('paginationContainer').classList.add('hidden');
    }
}

function updateFileList() {
    uploadedFilesDiv.classList.remove('hidden');
    fileList.innerHTML = '';
    uploadedFiles.forEach((file, index) => {
        const listItem = document.createElement('li');
        listItem.className = 'flex items-center justify-between p-2 bg-gray-50 rounded-md';
        const fileInfo = document.createElement('div');
        fileInfo.className = 'flex items-center';
        const fileIcon = document.createElement('i');
        const ext = file.name.split('.').pop().toLowerCase();
        if (['xlsx', 'xls'].includes(ext)) fileIcon.className = 'fa fa-file-excel-o text-success mr-2';
        else if (ext === 'csv') fileIcon.className = 'fa fa-file-text-o text-info mr-2';
        else if (['doc', 'docx'].includes(ext)) fileIcon.className = 'fa fa-file-word-o text-primary mr-2';
        else fileIcon.className = 'fa fa-file-o text-gray-400 mr-2';
        const fileName = document.createElement('span');
        fileName.className = 'text-sm text-gray-700 mr-2';
        fileName.textContent = file.name;
        const fileSize = document.createElement('span');
        fileSize.className = 'text-xs text-gray-500';
        fileSize.textContent = formatFileSize(file.size);
        fileInfo.appendChild(fileIcon);
        fileInfo.appendChild(fileName);
        fileInfo.appendChild(fileSize);
        const deleteBtn = document.createElement('button');
        deleteBtn.className = 'text-gray-400 hover:text-danger transition-colors';
        deleteBtn.innerHTML = '<i class="fa fa-times"></i>';
        deleteBtn.title = '删除文件';
        deleteBtn.addEventListener('click', () => {
            uploadedFiles.splice(index, 1);
            updateFileList();
            if (uploadedFiles.length === 0) {
                startReviewBtn.disabled = true;
                uploadedFilesDiv.classList.add('hidden');
            } else {
                // 只要有文件就启用按钮，没有活动规则时点击会显示弹窗
                startReviewBtn.disabled = false;
            }
        });
        listItem.appendChild(fileInfo);
        listItem.appendChild(deleteBtn);
        fileList.appendChild(listItem);
    });
}

// 将关键渲染函数暴露到全局，便于控制台调用与其他脚本使用
window.updateResultsTable = updateResultsTable;
window.displayReviewResults = displayReviewResults;
window.updateFileList = updateFileList;

function switchRuleTab(e) {
    const tabId = e.target.dataset.tab;
    ruleTabs.forEach(tab => { tab.classList.remove('text-primary', 'border-b-2', 'border-primary'); tab.classList.add('text-gray-500'); });
    e.target.classList.add('text-primary', 'border-b-2', 'border-primary');
    e.target.classList.remove('text-gray-500');
    ruleContents.forEach(content => content.classList.add('hidden'));
    document.getElementById(`${tabId}-content`).classList.remove('hidden');
}

function handleFilterChange(e) {
    const filter = e.target.dataset.filter;
    currentFilter = filter;
    // 重置分页状态
    currentPage = 1;
    updateResultsTable();
    filterBtns.forEach(btn => { btn.classList.remove('bg-primary', 'text-white'); btn.classList.add('bg-gray-200', 'text-gray-700'); });
    e.target.classList.add('bg-primary', 'text-white');
    e.target.classList.remove('bg-gray-200', 'text-gray-700');
}

function handleExportOptionChange(e) {
    const exportFilter = e.target.dataset.exportFilter;
    currentExportFilter = exportFilter;
    document.querySelectorAll('.export-option-btn').forEach(btn => { btn.classList.remove('bg-white', 'text-gray-700', 'shadow-sm'); btn.classList.add('bg-gray-100', 'text-gray-500'); });
    e.target.classList.add('bg-white', 'text-gray-700', 'shadow-sm');
    e.target.classList.remove('bg-gray-100', 'text-gray-500');
}

function toggleTheme() {
    document.body.classList.toggle('bg-gray-900');
    document.body.classList.toggle('text-white');
    const icon = themeToggle.querySelector('i');
    if (icon.classList.contains('fa-moon-o')) { icon.classList.remove('fa-moon-o'); icon.classList.add('fa-sun-o'); }
    else { icon.classList.remove('fa-sun-o'); icon.classList.add('fa-moon-o'); }
}

