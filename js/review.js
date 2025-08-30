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

// 合计计算和比对功能
window.calculateAndCompareTotal = function(records, originalText) {
    console.log('开始计算合计和比对...');
    
    // 1. 计算提取的活动分数总和
    const extractedScores = records.map(record => {
        const score = parseFloat(record['加分']) || 0;
        return {
            activity: record['所参加的活动及担任角色'],
            score: score,
            name: record['姓名']
        };
    });
    
    const calculatedTotal = extractedScores.reduce((sum, item) => sum + item.score, 0);
    console.log('计算得出的总分:', calculatedTotal);
    console.log('各活动分数:', extractedScores);
    
    // 2. 提取文档中的合计分数
    const documentTotal = extractDocumentTotal(originalText);
    console.log('文档中填写的合计:', documentTotal);
    
    // 3. 比对结果
    const comparisonResult = {
        calculatedTotal: calculatedTotal,
        documentTotal: documentTotal,
        isMatch: false,
        difference: 0,
        issues: [],
        extractedScores: extractedScores
    };
    
    if (documentTotal === null) {
        comparisonResult.issues.push('文档中未找到合计分数');
        comparisonResult.isMatch = false;
    } else if (Math.abs(calculatedTotal - documentTotal) < 0.01) {
        comparisonResult.isMatch = true;
        comparisonResult.issues.push('合计分数正确');
    } else {
        comparisonResult.isMatch = false;
        comparisonResult.difference = calculatedTotal - documentTotal;
        comparisonResult.issues.push(`合计分数不匹配：计算得出 ${calculatedTotal}，文档填写 ${documentTotal}，差异 ${comparisonResult.difference > 0 ? '+' : ''}${comparisonResult.difference.toFixed(1)}`);
    }
    
    return comparisonResult;
};

// 提取文档中的合计分数
function extractDocumentTotal(text) {
    if (!text) return null;
    
    // 多种合计匹配模式
    const totalPatterns = [
        // 合计：数字
        /合计[：:]\s*(\d+(?:\.\d+)?)/g,
        // 合计 数字
        /合计\s+(\d+(?:\.\d+)?)/g,
        // 总计：数字
        /总计[：:]\s*(\d+(?:\.\d+)?)/g,
        // 总分：数字
        /总分[：:]\s*(\d+(?:\.\d+)?)/g,
        // 总计得分：数字
        /总计得分[：:]\s*(\d+(?:\.\d+)?)/g,
        // 同意加分：数字
        /同意加分[：:]\s*(\d+(?:\.\d+)?)/g,
        // 表格中的合计行
        /合计\s*\t\s*(\d+(?:\.\d+)?)/g,
        // 合计行中的数字
        /合计.*?(\d+(?:\.\d+)?)/g
    ];
    
    let foundTotal = null;
    
    for (const pattern of totalPatterns) {
        const matches = text.match(pattern);
        if (matches && matches.length > 0) {
            // 取最后一个匹配的合计（通常是最准确的）
            const lastMatch = matches[matches.length - 1];
            const numberMatch = lastMatch.match(/(\d+(?:\.\d+)?)/);
            if (numberMatch) {
                foundTotal = parseFloat(numberMatch[1]);
                console.log('找到合计分数:', foundTotal, '匹配模式:', pattern);
                break;
            }
        }
    }
    
    return foundTotal;
}

// 在审查面板中显示合计比对结果
window.displayTotalComparison = function(comparisonResult) {
    if (!comparisonResult) return;
    
    const reviewPanel = document.getElementById('reviewResults');
    if (!reviewPanel) return;
    
    // 创建合计比对结果显示区域
    let totalComparisonHtml = `
        <div class="total-comparison-section" style="margin: 20px 0; padding: 15px; border: 2px solid #e9ecef; border-radius: 8px; background-color: #f8f9fa;">
            <h3 style="margin: 0 0 15px 0; color: #495057; border-bottom: 1px solid #dee2e6; padding-bottom: 10px;">
                📊 合计分数比对结果
            </h3>
    `;
    
    if (comparisonResult.isMatch) {
        totalComparisonHtml += `
            <div style="color: #155724; background-color: #d4edda; border: 1px solid #c3e6cb; padding: 10px; border-radius: 4px; margin-bottom: 15px;">
                ✅ 合计分数正确：${comparisonResult.calculatedTotal}
            </div>
        `;
    } else if (comparisonResult.documentTotal === null) {
        // 文档中没有找到合计分数
        totalComparisonHtml += `
            <div style="color: #856404; background-color: #fff3cd; border: 1px solid #ffeaa7; padding: 10px; border-radius: 4px; margin-bottom: 15px;">
                ⚠️ 文档中未找到合计分数：
                <br>• 计算得出：<strong>${comparisonResult.calculatedTotal}</strong>
                <br>• 建议在文档中添加合计分数以便核对
            </div>
        `;
    } else {
        // 合计分数不匹配
        totalComparisonHtml += `
            <div style="color: #721c24; background-color: #f8d7da; border: 1px solid #f5c6cb; padding: 10px; border-radius: 4px; margin-bottom: 15px;">
                ❌ 合计分数不匹配：
                <br>• 计算得出：<strong>${comparisonResult.calculatedTotal}</strong>
                <br>• 文档填写：<strong>${comparisonResult.documentTotal}</strong>
                <br>• 差异：<strong style="color: ${comparisonResult.difference > 0 ? '#dc3545' : '#28a745'}">${comparisonResult.difference > 0 ? '+' : ''}${comparisonResult.difference.toFixed(1)}</strong>
            </div>
        `;
    }
    
    // 显示各活动分数明细
    totalComparisonHtml += `
        <div style="margin-top: 15px;">
            <h4 style="margin: 0 0 10px 0; color: #495057;">活动分数明细：</h4>
            <div style="max-height: 200px; overflow-y: auto; border: 1px solid #dee2e6; border-radius: 4px; padding: 10px; background-color: white;">
    `;
    
    comparisonResult.extractedScores.forEach((item, index) => {
        totalComparisonHtml += `
            <div style="display: flex; justify-content: space-between; padding: 5px 0; border-bottom: 1px solid #f8f9fa;">
                <span style="flex: 1; font-size: 14px;">${index + 1}. ${item.activity}</span>
                <span style="font-weight: bold; color: #28a745; margin-left: 10px;">${item.score}</span>
            </div>
        `;
    });
    
    totalComparisonHtml += `
            </div>
            <div style="margin-top: 10px; padding: 10px; background-color: #e9ecef; border-radius: 4px; text-align: right;">
                <strong>总计：${comparisonResult.calculatedTotal}</strong>
            </div>
        </div>
    </div>
    `;
    
    // 将合计比对结果插入到审查面板的顶部
    reviewPanel.insertAdjacentHTML('afterbegin', totalComparisonHtml);
};

// 修改现有的审查函数，集成合计比对功能
window.reviewRecordsWithTotal = function(records, originalText) {
    // 先进行常规审查
    const reviewResults = records.map(record => reviewRecord(record));
    
    // 进行合计比对
    const totalComparison = calculateAndCompareTotal(records, originalText);
    
    // 如果合计分数不匹配，将问题添加到所有审查记录的问题描述中
    if (!totalComparison.isMatch && totalComparison.documentTotal !== null) {
        reviewResults.forEach(result => {
            if (result.issues && result.issues !== '通过') {
                result.issues += '; 合计分数有误';
            } else if (result.issues === '通过') {
                result.issues = '合计分数有误';
                result.passed = false;
            } else {
                result.issues = '合计分数有误';
                result.passed = false;
            }
        });
    }
    
    // 显示审查结果
    displayReviewResults(reviewResults);
    
    return {
        reviewResults: reviewResults,
        totalComparison: totalComparison
    };
};

