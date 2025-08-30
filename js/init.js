function initEventListeners() {
    console.log('开始初始化事件监听器...');
    console.log('startReviewBtn 元素:', startReviewBtn);
    
    browseBtn.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', handleFileSelect);
    dropArea.addEventListener('dragover', handleDragOver);
    dropArea.addEventListener('drop', handleDrop);
    
    if (startReviewBtn) {
        startReviewBtn.addEventListener('click', window.startReview);
        console.log('开始审查按钮事件监听器已绑定');
    } else {
        console.error('startReviewBtn 元素未找到');
    }
    exportResultsBtn.addEventListener('click', exportResults);
    clearResultsBtn.addEventListener('click', clearReviewResults);
    helpBtn.addEventListener('click', () => helpModal.classList.remove('hidden'));
    closeHelpModal.addEventListener('click', () => helpModal.classList.add('hidden'));
    closeActivityRuleModal.addEventListener('click', () => activityRuleModal.classList.add('hidden'));
    goToActivityRules.addEventListener('click', () => {
        activityRuleModal.classList.add('hidden');
        // 切换到活动规则选项卡
        ruleTabs.forEach(tab => { 
            tab.classList.remove('text-primary', 'border-b-2', 'border-primary'); 
            tab.classList.add('text-gray-500'); 
        });
        document.querySelector('[data-tab="activity"]').classList.add('text-primary', 'border-b-2', 'border-primary');
        document.querySelector('[data-tab="activity"]').classList.remove('text-gray-500');
        ruleContents.forEach(content => content.classList.add('hidden'));
        document.getElementById('activity-content').classList.remove('hidden');
        // 滚动到规则设置区域
        document.querySelector('#activity-content').scrollIntoView({ behavior: 'smooth' });
    });
    
    // 学号规则弹窗事件监听器
    closeStudentIdRuleModal.addEventListener('click', () => studentIdRuleModal.classList.add('hidden'));
    goToStudentIdRules.addEventListener('click', () => {
        studentIdRuleModal.classList.add('hidden');
        // 切换到学号规则选项卡
        ruleTabs.forEach(tab => { 
            tab.classList.remove('text-primary', 'border-b-2', 'border-primary'); 
            tab.classList.add('text-gray-500'); 
        });
        const studentIdTab = document.querySelector('[data-tab="student"]');
        if (studentIdTab) {
            studentIdTab.classList.add('text-primary', 'border-b-2', 'border-primary');
            studentIdTab.classList.remove('text-gray-500');
        }
        ruleContents.forEach(content => content.classList.add('hidden'));
        const studentIdContent = document.getElementById('student-content');
        if (studentIdContent) {
            studentIdContent.classList.remove('hidden');
            // 滚动到规则设置区域
            studentIdContent.scrollIntoView({ behavior: 'smooth' });
        }
    });
    
    // 部门规则弹窗事件监听器
    closeDepartmentRuleModal.addEventListener('click', () => departmentRuleModal.classList.add('hidden'));
    goToDepartmentRules.addEventListener('click', () => {
        departmentRuleModal.classList.add('hidden');
        // 切换到部门规则选项卡
        ruleTabs.forEach(tab => { 
            tab.classList.remove('text-primary', 'border-b-2', 'border-primary'); 
            tab.classList.add('text-gray-500'); 
        });
        const departmentTab = document.querySelector('[data-tab="department"]');
        if (departmentTab) {
            departmentTab.classList.add('text-primary', 'border-b-2', 'border-primary');
            departmentTab.classList.remove('text-gray-500');
        }
        ruleContents.forEach(content => content.classList.add('hidden'));
        const departmentContent = document.getElementById('department-content');
        if (departmentContent) {
            departmentContent.classList.remove('hidden');
            // 滚动到规则设置区域
            departmentContent.scrollIntoView({ behavior: 'smooth' });
        }
    });
    themeToggle.addEventListener('click', toggleTheme);
    ruleTabs.forEach(tab => tab.addEventListener('click', switchRuleTab));
    addActivityRule.addEventListener('click', handleAddActivityRule);
    document.getElementById('batchAddActivityRule').addEventListener('click', handleBatchAddActivityRule);
    saveStudentIdRule.addEventListener('click', handleAddStudentIdRule);
    addDepartmentRule.addEventListener('click', handleAddDepartmentRule);
    document.getElementById('batchAddDepartmentRule').addEventListener('click', handleBatchAddDepartmentRule);
    activitySelect.addEventListener('change', handleActivitySelectChange);
    departmentSelect.addEventListener('change', handleDepartmentSelectChange);
    filterBtns.forEach(btn => btn.addEventListener('click', handleFilterChange));
    document.querySelectorAll('.export-option-btn').forEach(btn => btn.addEventListener('click', handleExportOptionChange));
    helpModal.addEventListener('click', (e) => { if (e.target === helpModal) helpModal.classList.add('hidden'); });
    activityRuleModal.addEventListener('click', (e) => { if (e.target === activityRuleModal) activityRuleModal.classList.add('hidden'); });
    studentIdRuleModal.addEventListener('click', (e) => { if (e.target === studentIdRuleModal) studentIdRuleModal.classList.add('hidden'); });
    departmentRuleModal.addEventListener('click', (e) => { if (e.target === departmentRuleModal) departmentRuleModal.classList.add('hidden'); });
    window.addEventListener('dragenter', showDragOverlay, false);
    window.addEventListener('dragover', showDragOverlay, false);
    window.addEventListener('dragleave', maybeHideDragOverlay, false);
    window.addEventListener('drop', handleGlobalDrop, false);

    // 查询事件
    if (applyQueryBtn) applyQueryBtn.addEventListener('click', () => {
        currentQuery.department = (queryDepartment && queryDepartment.value) || '';
        currentQuery.studentPrefix = (queryStudentPrefix && queryStudentPrefix.value.trim()) || '';
        updateResultsTable();
    });
    if (resetQueryBtn) resetQueryBtn.addEventListener('click', () => {
        if (queryDepartment) queryDepartment.value = '';
        if (queryStudentPrefix) queryStudentPrefix.value = '';
        currentQuery = { department: '', studentPrefix: '' };
        updateResultsTable();
    });
    
    // 分页事件监听器
    document.getElementById('firstPageBtn').addEventListener('click', () => goToPage(1));
    document.getElementById('prevPageBtn').addEventListener('click', () => goToPage(currentPage - 1));
    document.getElementById('nextPageBtn').addEventListener('click', () => goToPage(currentPage + 1));
    document.getElementById('lastPageBtn').addEventListener('click', () => goToPage(totalPages));
    document.getElementById('pageSizeSelect').addEventListener('change', (e) => changePageSize(e.target.value));
    document.getElementById('currentPageInput').addEventListener('keypress', (e) => {
        if (e.key === 'Enter') {
            const page = parseInt(e.target.value);
            if (!isNaN(page)) goToPage(page);
        }
    });
}

window.startReview = function() {
    console.log('startReview 函数被调用，activityRules.length:', activityRules.length, 'studentIdRules.length:', studentIdRules.length, 'departmentRules.length:', departmentRules.length, 'uploadedFiles.length:', uploadedFiles.length);
    
    // 首先检查是否上传了文件
    if (uploadedFiles.length === 0) {
        console.log('没有上传文件，显示提示');
        showNotification('请先上传文件后再进行审查', 'warning');
        return;
    }
    
    // 检查是否添加了必要的规则
    if (activityRules.length === 0) {
        console.log('没有活动规则，显示弹窗');
        if (activityRuleModal) {
            activityRuleModal.classList.remove('hidden');
            console.log('活动规则弹窗已显示');
        } else {
            console.error('activityRuleModal 元素未找到');
        }
        return;
    }
    
    if (studentIdRules.length === 0) {
        console.log('没有学号规则，显示弹窗');
        if (studentIdRuleModal) {
            studentIdRuleModal.classList.remove('hidden');
            console.log('学号规则弹窗已显示');
        } else {
            console.error('studentIdRuleModal 元素未找到');
        }
        return;
    }
    
    if (departmentRules.length === 0) {
        console.log('没有部门规则，显示弹窗');
        if (departmentRuleModal) {
            departmentRuleModal.classList.remove('hidden');
            console.log('部门规则弹窗已显示');
        } else {
            console.error('departmentRuleModal 元素未找到');
        }
        return;
    }
    
    reviewResults = [];
    resultsTable.innerHTML = '';
    reviewStatus.classList.remove('hidden');
    resultFilter.classList.add('hidden');
    resultStats.classList.add('hidden');
    document.getElementById('exportOptions').classList.add('hidden');
    exportResultsBtn.disabled = true;
    clearResultsBtn.classList.add('hidden');
    progressBar.style.width = '0%';
    progressPercentage.textContent = '0%';
    let progress = 0;
    const interval = setInterval(() => {
        progress += 10;
        progressBar.style.width = `${progress}%`;
        progressPercentage.textContent = `${progress}%`;
        if (progress >= 100) {
            clearInterval(interval);
            setTimeout(() => {
                Promise.all(uploadedFiles.map(processFile)).then(() => {
                    reviewStatus.classList.add('hidden');
                    if (reviewResults.length === 0) {
                        showNotification('未能从上传的文件中提取有效记录，请检查文件格式是否正确', 'error');
                        showNotification('请按F12打开控制台查看详细的解析日志', 'info');
                    } else {
                        displayReviewResults();
                        exportResultsBtn.disabled = false;
                        clearResultsBtn.classList.remove('hidden');
                        resultFilter.classList.remove('hidden');
                        resultStats.classList.remove('hidden');
                    }
                });
            }, 500);
        }
    }, 200);
}

function initApp() {
    console.log('initApp 开始执行...');
    console.log('activityRules 初始值:', activityRules);
    
    initEventListeners();
    
    activityRules = [
       
    ];
    console.log('activityRules 设置后:', activityRules);
    
    updateActivityRulesTable();
    updateActivitySelect();
    departmentRules = [ ];
    updateDepartmentRulesTable();
    updateDepartmentSelect();
    // 初始化查询的部门下拉（与规则部门一致）
    updateQueryDepartmentSelect();
    studentIdRules = [];
    updateStudentIdRulesTable();
    
    console.log('initApp 执行完成');
    
    // 为了测试弹窗功能，临时启用开始审查按钮
    if (startReviewBtn) {
        startReviewBtn.disabled = false;
        console.log('开始审查按钮已启用（用于测试弹窗）');
    }
}

window.addEventListener('DOMContentLoaded', initApp);

