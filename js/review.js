window.reviewRecord = function(record) {
    const rawScore = record['加分'];
    const isScoreMissing = rawScore === undefined || rawScore === null || String(rawScore).trim() === '';
    const activityName = record['所参加的活动及担任角色'] || '未知';
    
    // 提取年份信息（匹配4位数字的年份）
    let year = '';
    const yearMatch = activityName.match(/\b(20\d{2}|19\d{2})\b/);
    if (yearMatch && yearMatch[1]) {
        year = yearMatch[1];
    }
    
    const result = {
        name: record['姓名'] || '未知',
        studentId: record['学号'] || '未知',
        department: record['所在部门'] || '未知',
        activityName: activityName,
        year: year,  // 添加年份字段
        score: isScoreMissing ? 0 : (parseFloat(record['加分']) || 0),
        passed: true,
        issues: ''
    };
    result._scoreMissing = isScoreMissing;
    // 部门统一规范化
    const normalizedDept = normalizeDepartment(result.department);
    result.department = normalizedDept && normalizedDept.length > 0 ? normalizedDept : '未知';
    const issues = [];
    
    // 首先检查是否添加了必要的规则
    if (activityRules.length === 0) {
        issues.push('活动规则未添加');
        result.passed = false;
        result.issues = issues.join('; ');
        return result;
    }
    if (studentIdRules.length === 0) {
        issues.push('学号规则未添加');
        result.passed = false;
        result.issues = issues.join('; ');
        return result;
    }
    if (departmentRules.length === 0) {
        issues.push('部门规则未添加');
        result.passed = false;
        result.issues = issues.join('; ');
        return result;
    }
    if (result.studentId === '未知') { issues.push('缺少学号信息'); result.passed = false; }
    else if (studentIdRules.length > 0) {
        const prefixRules = studentIdRules.filter(r => r.type === 'prefix');
        const lengthRules = studentIdRules.filter(r => r.type === 'length');
        let prefixValid = prefixRules.length === 0;
        if (prefixRules.length > 0) {
            prefixValid = prefixRules.some(rule => result.studentId.startsWith(rule.value));
            if (!prefixValid) { issues.push(`学号格式不符合规则: ${result.studentId}`); result.passed = false; }
        }
        let lengthValid = lengthRules.length === 0;
        if (lengthRules.length > 0) {
            lengthValid = lengthRules.some(rule => result.studentId.length === parseInt(rule.value));
            if (!lengthValid) { issues.push('学号位数不符合'); result.passed = false; }
        }
    }
    if (!result.department || result.department === '未知') { issues.push('未填写部门'); result.passed = false; }
    else if (departmentRules.length > 0) {
        let departmentMatched = false;
        for (const dept of departmentRules) {
            const nd = normalizeDepartment(dept);
            if (!nd) continue;
            if (result.department.includes(nd) || nd.includes(result.department)) { departmentMatched = true; break; }
        }
        if (!departmentMatched) { issues.push(`部门名称未在规则列表中: ${result.department}`); result.passed = false; }
    }
    let matchedRuleRef = null;
    if (result.activityName !== '未知') {
        const normalizedActivity = normalizeActivityName(result.activityName);
        matchedRuleRef = activityRules.find(rule => {
            const rn = normalizeActivityName(rule.name);
            return normalizedActivity.includes(rn) || rn.includes(normalizedActivity);
        });
        if (matchedRuleRef) {
            if (result._scoreMissing) { issues.push('<span class="text-red-600">未填写活动分数</span>'); result.passed = false; }
            else if (Math.abs(result.score - matchedRuleRef.score) > 0.01) { issues.push(`活动分数不匹配，应为 ${matchedRuleRef.score}，实际为 ${result.score}`); result.passed = false; }
        } else { issues.push(`活动名称未在规则列表中: ${result.activityName}`); result.passed = false; }
    }
    if (result.name === '未知') { issues.push('缺少姓名信息'); result.passed = false; }
    if (result.activityName === '未知') { issues.push('缺少活动名称信息'); result.passed = false; }
    // 只有在有活动规则的情况下才检查分数相关问题
    if (activityRules.length > 0) {
        // 文档未填写分数：直接提示红色文案
        if (result._scoreMissing && !matchedRuleRef) { issues.push('<span class="text-red-600">未填写活动分数</span>'); result.passed = false; }
        // 未匹配任何活动规则且分数无效（0或非数）
        if (!matchedRuleRef && !result._scoreMissing && (result.score === 0 || isNaN(result.score))) { issues.push('活动分数无效或未设置'); result.passed = false; }
    }
    result.issues = issues.join('; ');
    if (issues.length === 0) { result.passed = true; result.issues = '通过'; }
    return result;
}

function clearReviewResults() {
    if (reviewResults.length === 0) { showNotification('当前没有可清空的审查结果', 'info'); return; }
    const confirmed = confirm('确定要清空所有审查结果吗？此操作不可撤销。');
    if (!confirmed) return;
    reviewResults = [];
    resultsTable.innerHTML = '';
    totalRecords.textContent = '0';
    passedRecords.textContent = '0';
    failedRecords.textContent = '0';
    exportResultsBtn.disabled = true;
    document.getElementById('exportOptions').classList.add('hidden');
    resultFilter.classList.add('hidden');
    resultStats.classList.add('hidden');
    clearResultsBtn.classList.add('hidden');
    const emptyRow = document.createElement('tr');
    emptyRow.innerHTML = `
        <td colspan="7" class="px-4 py-8 text-center text-gray-500">
            <i class="fa fa-info-circle text-2xl mb-2 block opacity-50"></i>
            没有记录
        </td>
    `;
    resultsTable.appendChild(emptyRow);
    showNotification('已清空所有审查结果', 'success');
}

function exportResults() {
    if (reviewResults.length === 0) { showNotification('没有可导出的结果', 'error'); return; }
    let filteredData = reviewResults; let exportType = '全部';
    if (currentExportFilter === 'passed') { filteredData = reviewResults.filter(r => r.passed); exportType = '通过'; }
    else if (currentExportFilter === 'failed') { filteredData = reviewResults.filter(r => !r.passed); exportType = '不通过'; }
    if (filteredData.length === 0) { showNotification(`没有${exportType}的记录可导出`, 'error'); return; }
    const exportData = filteredData.map(result => ({ '姓名': result.name, '学号': result.studentId, '部门': result.department, '活动名称': result.activityName, '分数': result.score, '审查结果': result.passed ? '通过' : '未通过', '问题描述': result.issues }));
    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, '审查结果');
    const fileName = `学生活动积分审查结果_${exportType}_${new Date().toLocaleDateString()}.xlsx`;
    XLSX.writeFile(workbook, fileName);
    showNotification(`${exportType}结果导出成功，共${filteredData.length}条记录`, 'success');
}

