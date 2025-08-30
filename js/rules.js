function updateActivityRulesTable() {
    activityRulesTable.innerHTML = '';
    activityRules.forEach((rule, index) => {
        const row = document.createElement('tr');
        row.className = index % 2 === 0 ? 'bg-white' : 'bg-gray-50';
        row.innerHTML = `
            <td class="px-4 py-2 text-sm text-gray-700">${rule.name}</td>
            <td class="px-4 py-2 text-sm text-gray-700">${rule.score}</td>
            <td class="px-4 py-2 text-sm">
                <button class="text-danger hover:text-red-700 transition-colors delete-activity-rule" data-index="${index}">
                    <i class="fa fa-trash"></i> 删除
                </button>
            </td>
        `;
        activityRulesTable.appendChild(row);
    });
    document.querySelectorAll('.delete-activity-rule').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const index = parseInt(e.target.closest('button').dataset.index);
            activityRules.splice(index, 1);
            updateActivityRulesTable();
            updateActivitySelect();
            
                // 更新开始审查按钮状态 - 只要有文件就启用按钮
    if (uploadedFiles.length > 0) {
        startReviewBtn.disabled = false;
    }
            
            showNotification('活动规则已删除', 'success');
        });
    });
    updateActivitySelect();
}

function handleAddActivityRule() {
    const name = normalizeActivityName(activityName.value.trim());
    const score = parseFloat(activityScore.value);
    if (!name) { showNotification('请输入活动名称', 'error'); return; }
    if (isNaN(score) || score < 0) { showNotification('请输入有效的分数', 'error'); return; }
    const existingRule = activityRules.find(rule => rule.name === name);
    if (existingRule) { existingRule.score = score; showNotification('活动规则已更新', 'success'); }
    else { activityRules.push({ name, score }); showNotification('活动规则已添加', 'success'); }
    activityName.value = ''; activityScore.value = ''; activitySelect.value = '';
    updateActivityRulesTable(); updateActivitySelect();
    
    // 更新开始审查按钮状态
    if (uploadedFiles.length > 0) {
        startReviewBtn.disabled = false;
    }
}

function handleBatchAddActivityRule() {
    const batchInput = document.getElementById('batchActivityInput').value.trim();
    if (!batchInput) { showNotification('请输入要批量添加的活动信息', 'error'); return; }
    const lines = batchInput.split('\n'); let successCount = 0; let errorCount = 0; const errors = [];
    lines.forEach((line, index) => {
        line = line.trim(); if (!line) return;
        const parts = line.split(','); if (parts.length !== 2) { errorCount++; errors.push(`第${index + 1}行格式错误：${line}`); return; }
        const name = normalizeActivityName(parts[0].trim()); const score = parseFloat(parts[1].trim());
        if (!name) { errorCount++; errors.push(`第${index + 1}行活动名称为空`); return; }
        if (isNaN(score) || score < 0) { errorCount++; errors.push(`第${index + 1}行分数无效：${parts[1]}`); return; }
        const existingRule = activityRules.find(rule => rule.name === name);
        if (existingRule) { existingRule.score = score; successCount++; } else { activityRules.push({ name, score }); successCount++; }
    });
    document.getElementById('batchActivityInput').value = ''; activitySelect.value = '';
    updateActivityRulesTable(); updateActivitySelect();
    
    // 更新开始审查按钮状态
    if (successCount > 0 && uploadedFiles.length > 0) {
        startReviewBtn.disabled = false;
    }
    
    if (successCount > 0 && errorCount === 0) showNotification(`成功批量添加 ${successCount} 个活动规则`, 'success');
    else if (successCount > 0 && errorCount > 0) showNotification(`成功添加 ${successCount} 个，失败 ${errorCount} 个。请检查格式。`, 'error');
    else showNotification('批量添加失败，请检查输入格式', 'error');
}

function handleAddStudentIdRule() {
    const ruleType = document.getElementById('studentIdRuleType').value;
    const value = studentIdPatternInput.value.trim();
    if (!ruleType) { showNotification('请选择规则类型', 'error'); return; }
    if (!value) { showNotification('请输入规则值', 'error'); return; }
    if (ruleType === 'prefix') { if (!/^[a-zA-Z0-9]+$/.test(value)) { showNotification('开头规则只能包含字母和数字', 'error'); return; } }
    else if (ruleType === 'length') { const length = parseInt(value); if (isNaN(length) || length <= 0 || length > 20) { showNotification('请输入1-20之间的有效位数', 'error'); return; } }
    const existingRule = studentIdRules.find(rule => rule.type === ruleType && rule.value === value);
    if (existingRule) { showNotification('该规则已存在', 'error'); return; }
    const newRule = { type: ruleType, value: value, description: ruleType === 'prefix' ? `以${value}开头` : '十位学号' };
    studentIdRules.push(newRule);
    document.getElementById('studentIdRuleType').value = '';
    studentIdPatternInput.value = '';
    updateStudentIdRulesTable();
    showNotification('学号规则已添加', 'success');
}

function handleAddDepartmentRule() {
    const name = departmentName.value.trim();
    if (!name) { showNotification('请输入部门名称', 'error'); return; }
    if (departmentRules.includes(name)) { showNotification('部门已存在', 'error'); return; }
    departmentRules.push(name);
    departmentName.value = ''; departmentSelect.value = '';
    updateDepartmentRulesTable(); updateDepartmentSelect();
    showNotification('部门规则已添加', 'success');
}

function handleBatchAddDepartmentRule() {
    const batchInput = document.getElementById('batchDepartmentInput').value.trim();
    if (!batchInput) { 
        showNotification('请输入要批量添加的部门信息', 'error'); 
        return; 
    }
    
    const departments = batchInput.split('\n').map(dept => dept.trim()).filter(dept => dept.length > 0);
    if (departments.length === 0) { 
        showNotification('没有有效的部门名称', 'error'); 
        return; 
    }
    
    let successCount = 0;
    let errorCount = 0;
    const existingDepartments = new Set(departmentRules);
    
    departments.forEach(dept => {
        if (dept && !existingDepartments.has(dept)) {
            departmentRules.push(dept);
            existingDepartments.add(dept);
            successCount++;
        } else if (existingDepartments.has(dept)) {
            errorCount++;
        }
    });
    
    // 清空输入框
    document.getElementById('batchDepartmentInput').value = '';
    
    // 更新界面
    updateDepartmentRulesTable();
    updateDepartmentSelect();
    
    // 显示结果通知
    if (successCount > 0 && errorCount === 0) {
        showNotification(`成功批量添加 ${successCount} 个部门规则`, 'success');
    } else if (successCount > 0 && errorCount > 0) {
        showNotification(`成功添加 ${successCount} 个，失败 ${errorCount} 个。请检查重复项。`, 'warning');
    } else {
        showNotification('批量添加失败，所有部门都已存在', 'error');
    }
}

function updateStudentIdRulesTable() {
    const studentIdRulesTable = document.getElementById('studentIdRulesTable');
    studentIdRulesTable.innerHTML = '';
    studentIdRules.forEach((rule, index) => {
        const row = document.createElement('tr');
        row.className = index % 2 === 0 ? 'bg-white' : 'bg-gray-50';
        const typeText = rule.type === 'prefix' ? '开头规则' : '位数规则';
        const valueText = rule.type === 'prefix' ? `以${rule.value}开头` : `${rule.value}位`;
        row.innerHTML = `
            <td class="px-2 py-2 text-sm text-gray-700">${typeText}</td>
            <td class="px-2 py-2 text-sm text-gray-700">${valueText}</td>
            <td class="px-2 py-2 text-sm text-right">
                <button class="text-danger hover:text-red-700 transition-colors delete-student-id-rule" data-index="${index}">
                    <i class="fa fa-trash"></i> 删除
                </button>
            </td>
        `;
        studentIdRulesTable.appendChild(row);
    });
    document.querySelectorAll('.delete-student-id-rule').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const index = parseInt(e.target.closest('button').dataset.index);
            studentIdRules.splice(index, 1);
            updateStudentIdRulesTable();
            showNotification('学号规则已删除', 'success');
        });
    });
}

function updateDepartmentRulesTable() {
    departmentRulesTable.innerHTML = '';
    const headerRow = document.createElement('tr');
    headerRow.className = 'bg-gray-50';
    headerRow.innerHTML = `
        <th class="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">部门名称</th>
        <th class="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">操作</th>
    `;
    departmentRulesTable.appendChild(headerRow);
    departmentRules.forEach((rule, index) => {
        const row = document.createElement('tr'); row.className = index % 2 === 0 ? 'bg-white' : 'bg-gray-50';
        row.innerHTML = `
            <td class="px-4 py-2 text-sm text-gray-700">${rule}</td>
            <td class="px-4 py-2 text-sm">
                <button class="text-danger hover:text-red-700 transition-colors delete-department-rule" data-index="${index}">
                    <i class="fa fa-trash"></i> 删除
                </button>
            </td>
        `;
        departmentRulesTable.appendChild(row);
    });
    document.querySelectorAll('.delete-department-rule').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const index = parseInt(e.target.closest('button').dataset.index);
            departmentRules.splice(index, 1);
            updateDepartmentRulesTable();
            updateDepartmentSelect();
            showNotification('部门规则已删除', 'success');
        });
    });
}

function updateActivitySelect() {
    activitySelect.innerHTML = '';
    const defaultOption = document.createElement('option'); defaultOption.value = ''; defaultOption.textContent = '请选择活动';
    activitySelect.appendChild(defaultOption);
    activityRules.forEach(rule => { const option = document.createElement('option'); option.value = rule.name; option.textContent = `${rule.name} (${rule.score}分)`; activitySelect.appendChild(option); });
}

function handleActivitySelectChange() {
    const selectedActivity = activitySelect.value;
    if (!selectedActivity) { activityName.value = ''; activityScore.value = ''; return; }
    const rule = activityRules.find(r => normalizeActivityName(r.name) === normalizeActivityName(selectedActivity));
    if (rule) { activityName.value = rule.name; activityScore.value = rule.score; }
}

function updateDepartmentSelect() {
    departmentSelect.innerHTML = '';
    const defaultOption = document.createElement('option'); defaultOption.value = ''; defaultOption.textContent = '请选择部门';
    departmentSelect.appendChild(defaultOption);
    departmentRules.forEach(dept => { const option = document.createElement('option'); option.value = dept; option.textContent = dept; departmentSelect.appendChild(option); });
    // 同步更新查询的部门下拉
    updateQueryDepartmentSelect();
}

// 更新查询的部门下拉选项
function updateQueryDepartmentSelect() {
    const queryDeptSelect = document.getElementById('queryDepartment');
    if (queryDeptSelect) {
        const currentValue = queryDeptSelect.value; // 保存当前选中值
        queryDeptSelect.innerHTML = '<option value="">全部部门</option>' +
            departmentRules.map(d => `<option value="${d}">${d}</option>`).join('');
        // 如果之前选中的部门仍然存在，则恢复选中状态
        if (currentValue && departmentRules.includes(currentValue)) {
            queryDeptSelect.value = currentValue;
        }
    }
}

function handleDepartmentSelectChange() {
    const selectedDepartment = departmentSelect.value;
    if (!selectedDepartment) { departmentName.value = ''; return; }
    departmentName.value = selectedDepartment;
}

