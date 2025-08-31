// 审查结果相关功能

// 审查单条记录的核心函数
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
    
    if (result.studentId === '未知') { 
        issues.push('缺少学号信息'); 
        result.passed = false; 
    } else if (studentIdRules.length > 0) {
        const prefixRules = studentIdRules.filter(r => r.type === 'prefix');
        const lengthRules = studentIdRules.filter(r => r.type === 'length');
        let prefixValid = prefixRules.length === 0;
        if (prefixRules.length > 0) {
            prefixValid = prefixRules.some(rule => result.studentId.startsWith(rule.value));
            if (!prefixValid) { 
                issues.push(`学号格式不符合规则: ${result.studentId}`); 
                result.passed = false; 
            }
        }
        let lengthValid = lengthRules.length === 0;
        if (lengthRules.length > 0) {
            lengthValid = lengthRules.some(rule => result.studentId.length === parseInt(rule.value));
            if (!lengthValid) { 
                issues.push('学号位数不符合'); 
                result.passed = false; 
            }
        }
    }
    
    if (!result.department || result.department === '未知') { 
        issues.push('未填写部门'); 
        result.passed = false; 
    } else if (departmentRules.length > 0) {
        let departmentMatched = false;
        for (const dept of departmentRules) {
            const nd = normalizeDepartment(dept);
            if (!nd) continue;
            if (result.department.includes(nd) || nd.includes(result.department)) { 
                departmentMatched = true; 
                break; 
            }
        }
        if (!departmentMatched) { 
            issues.push(`部门名称未在规则列表中: ${result.department}`); 
            result.passed = false; 
        }
    }
    
    let matchedRuleRef = null;
    let cadreMatched = false;
    
    if (result.activityName !== '未知') {
        const normalizedActivity = normalizeActivityName(result.activityName);
        
        // 检查是否匹配活动规则
        matchedRuleRef = activityRules.find(rule => {
            const rn = normalizeActivityName(rule.name);
            return normalizedActivity.includes(rn) || rn.includes(normalizedActivity);
        });
        
        // 如果没有精确匹配，尝试相似度匹配
        if (!matchedRuleRef && activityRules.length > 0) {
            let bestSimilarity = null;
            let bestRule = null;
            
            console.log('开始相似度匹配，活动名称:', result.activityName);
            console.log('当前相似度设置:', similaritySettings);
            
            for (const rule of activityRules) {
                const similarity = calculateActivitySimilarity(result.activityName, rule.name);
                console.log(`比较 "${result.activityName}" 与 "${rule.name}":`, similarity);
                
                if (similarity.passed && (!bestSimilarity || similarity.fieldMatchCount > bestSimilarity.fieldMatchCount)) {
                    bestSimilarity = similarity;
                    bestRule = rule;
                    console.log('找到更好的匹配:', rule.name, similarity);
                }
            }
            
            if (bestSimilarity && bestRule) {
                matchedRuleRef = bestRule;
                console.log('相似度匹配成功:', bestRule.name, bestSimilarity);
                // 相似度匹配成功，不显示详细信息
            } else {
                console.log('相似度匹配失败，未找到符合条件的规则');
            }
        }
        
        // 检查是否匹配干部职位规则
        if (cadreRules.length > 0) {
            cadreMatched = cadreRules.some(cadre => 
                result.activityName.includes(cadre)
            );
        }
        
        if (matchedRuleRef) {
            if (result._scoreMissing) { 
                issues.push('<span class="text-red-600">未填写活动分数</span>'); 
                result.passed = false; 
            } else if (Math.abs(result.score - matchedRuleRef.score) > 0.01) { 
                issues.push(`活动分数不匹配，应为 ${matchedRuleRef.score}，实际为 ${result.score}`); 
                result.passed = false; 
            }
        } else if (cadreMatched) {
            // 如果匹配干部职位规则，则通过
            result.passed = true;
            // 找到匹配的干部职位
            const matchedCadre = cadreRules.find(cadre => 
                result.activityName.includes(cadre)
            );
            if (matchedCadre) {
                issues.push(`匹配干部职位规则: ${matchedCadre}`);
            }
        } else { 
            issues.push(`活动名称未在规则列表中: ${result.activityName}`); 
            result.passed = false; 
        }
    }
    
    if (result.name === '未知') { 
        issues.push('缺少姓名信息'); 
        result.passed = false; 
    }
    if (result.activityName === '未知') { 
        issues.push('缺少活动名称信息'); 
        result.passed = false; 
    }
    
    // 只有在有活动规则的情况下才检查分数相关问题
    if (activityRules.length > 0) {
        // 文档未填写分数：直接提示红色文案
        if (result._scoreMissing && !matchedRuleRef) { 
            issues.push('<span class="text-red-600">未填写活动分数</span>'); 
            result.passed = false; 
        }
        // 未匹配任何活动规则且分数无效（0或非数）
        if (!matchedRuleRef && !result._scoreMissing && (result.score === 0 || isNaN(result.score))) { 
            issues.push('活动分数无效或未设置'); 
            result.passed = false; 
        }
    }
    
    result.issues = issues.join('; ');
    if (issues.length === 0) { 
        result.passed = true; 
        result.issues = '通过'; 
    }
    
    return result;
};

// 显示审查结果
function displayReviewResults(results) {
    const resultsTable = document.getElementById('resultsTable');
    const resultStats = document.getElementById('resultStats');
    const resultFilter = document.getElementById('resultFilter');
    const clearResultsBtn = document.getElementById('clearResultsBtn');
    
    if (!resultsTable || !resultStats || !resultFilter) {
        console.error('必要的DOM元素未找到');
        return;
    }
    
    // 清空表格
    resultsTable.innerHTML = '';
    
    // 统计结果
    const totalRecords = results.length;
    const passedRecords = results.filter(r => r.passed).length;
    const failedRecords = results.filter(r => !r.passed).length;
    
    // 更新统计信息
    document.getElementById('totalRecords').textContent = totalRecords;
    document.getElementById('passedRecords').textContent = passedRecords;
    document.getElementById('failedRecords').textContent = failedRecords;
    
    // 显示结果
    results.forEach((result, index) => {
        const row = document.createElement('tr');
        row.className = index % 2 === 0 ? 'bg-white' : 'bg-gray-50';
        
        // 部门显示逻辑
        let deptDisplay = result.department;
        if (result.deptStatus === 'failed') {
            deptDisplay = `<span class="dept-failed">${result.department}</span>`;
        } else if (result.deptStatus === 'passed') {
            deptDisplay = `<span class="dept-passed">${result.department}</span>`;
        } else if (result.deptStatus === 'warning') {
            deptDisplay = `<span class="dept-warning">${result.department}</span>`;
        } else if (result.deptStatus === 'info') {
            deptDisplay = `<span class="dept-info">${result.department}</span>`;
        }
        
        row.innerHTML = `
            <td class="px-4 py-2 text-sm text-gray-700">${index + 1}</td>
            <td class="px-4 py-2 text-sm text-gray-700">${result.name || '未知'}</td>
            <td class="px-4 py-2 text-sm text-gray-700">${result.studentId || '未知'}</td>
            <td class="px-4 py-2 text-sm text-gray-700">${deptDisplay}</td>
            <td class="px-4 py-2 text-sm text-gray-700">${result.activityName || '未知'}</td>
            <td class="px-4 py-2 text-sm text-gray-700">${result.score || '未知'}</td>
            <td class="px-4 py-2 text-sm text-gray-700">
                <span class="${result.passed ? 'text-green-600' : 'text-red-600'} font-medium">
                    ${result.passed ? '通过' : '未通过'}
                </span>
            </td>
            <td class="px-4 py-2 text-sm text-gray-700 max-w-xs">
                <div class="whitespace-pre-wrap">${result.issues || '无'}</div>
            </td>
        `;
        resultsTable.appendChild(row);
    });
    
    // 显示相关界面元素
    resultStats.classList.remove('hidden');
    resultFilter.classList.remove('hidden');
    clearResultsBtn.classList.remove('hidden');
    
    // 更新分页
    updatePagination(totalRecords);
    
    // 隐藏审查状态
    document.getElementById('reviewStatus').classList.add('hidden');
    
    // 导出按钮样式更新
    if (exportDropdownBtn) {
        exportDropdownBtn.classList.remove('no-data');
    }
    
    console.log('审查结果已显示，共', totalRecords, '条记录');
}

// 清空审查结果
function clearReviewResults() {
    if (reviewResults.length === 0) {
        showNotification('当前没有可清空的审查结果', 'info');
        return;
    }
    
    reviewResults = [];
    document.getElementById('resultsTable').innerHTML = '';
    document.getElementById('resultStats').classList.add('hidden');
    document.getElementById('resultFilter').classList.add('hidden');
    if (exportDropdown) exportDropdown.classList.add('hidden');
    document.getElementById('clearResultsBtn').classList.add('hidden');
    document.getElementById('paginationContainer').classList.add('hidden');
    
    // 导出按钮样式更新
    if (exportDropdownBtn) {
        exportDropdownBtn.classList.add('no-data');
    }
    
    showNotification('已清空所有审查结果', 'success');
}

// 导出审查结果
function exportReviewResults(exportFilter = 'all') {
    if (reviewResults.length === 0) {
        showNotification('当前没有审查结果可导出，请先上传文件并开始审查', 'info');
        return;
    }
    
    let filteredData = [];
    let exportType = '';
    
    if (exportFilter === 'all') { filteredData = reviewResults; exportType = '全部'; }
    else if (exportFilter === 'passed') { filteredData = reviewResults.filter(r => r.passed); exportType = '通过'; }
    else if (exportFilter === 'failed') { filteredData = reviewResults.filter(r => !r.passed); exportType = '不通过'; }
    
    if (filteredData.length === 0) { showNotification(`没有${exportType}的记录可导出`, 'error'); return; }
    
    const exportData = filteredData.map(result => ({
        '姓名': result.name,
        '学号': result.studentId,
        '部门': result.department,
        '活动名称': result.activityName,
        '分数': result.score,
        '审查结果': result.passed ? '通过' : '未通过',
        '问题描述': result.issues
    }));
    
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
    console.log('输入记录数量:', records.length);
    console.log('所有记录:', records);
    
    // 1. 计算提取的活动分数总和（去重处理）
    const uniqueActivities = new Map();
    
    records.forEach(record => {
        const key = `${record['姓名']}-${record['所参加的活动及担任角色']}-${record['加分']}`;
        if (!uniqueActivities.has(key)) {
            uniqueActivities.set(key, {
                activity: record['所参加的活动及担任角色'],
                score: parseFloat(record['加分']) || 0,
                name: record['姓名']
            });
        } else {
            console.log('跳过重复活动:', record['所参加的活动及担任角色']);
        }
    });
    
    const extractedScores = Array.from(uniqueActivities.values());
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
        comparisonResult.issues.push(`合计分数有误：应为 ${calculatedTotal}，实际为 ${documentTotal}，差异 ${comparisonResult.difference > 0 ? '+' : ''}${comparisonResult.difference.toFixed(1)}`);
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

// 修改现有的审查函数，集成合计比对功能
window.reviewRecordsWithTotal = function(records, originalText) {
    // 先进行常规审查
    const reviewResults = records.map(record => reviewRecord(record));
    
    // 检查是否所有活动都通过审查
    const allActivitiesPassed = reviewResults.every(result => result.passed);
    
    console.log('所有活动是否都通过:', allActivitiesPassed);
    console.log('审查结果详情:', reviewResults.map(r => ({ activity: r.activityName, passed: r.passed, issues: r.issues })));
    
    let totalComparison = null;
    
    // 只有当所有活动都通过时，才进行合计分数比对
    if (allActivitiesPassed) {
        // 进行合计比对
        totalComparison = calculateAndCompareTotal(records, originalText);
        
        // 如果合计分数不匹配，将问题添加到所有审查记录的问题描述中
        if (!totalComparison.isMatch && totalComparison.documentTotal !== null) {
            const totalErrorMsg = `合计分数有误：应为 ${totalComparison.calculatedTotal}，实际为 ${totalComparison.documentTotal}，差异 ${totalComparison.difference > 0 ? '+' : ''}${totalComparison.difference.toFixed(1)}`;
            reviewResults.forEach(result => {
                if (result.issues && result.issues !== '通过') {
                    result.issues += '; ' + totalErrorMsg;
                } else if (result.issues === '通过') {
                    result.issues = totalErrorMsg;
                    result.passed = false;
                } else {
                    result.issues = totalErrorMsg;
                    result.passed = false;
                }
            });
        }
    } else {
        console.log('存在未通过的活动，跳过合计分数比对');
    }
    
    // 显示审查结果
    displayReviewResults(reviewResults);
    
    return {
        reviewResults: reviewResults,
        totalComparison: totalComparison
    };
};

// 绑定事件
document.addEventListener('DOMContentLoaded', function() {
    // 清空结果按钮
    const clearResultsBtn = document.getElementById('clearResultsBtn');
    if (clearResultsBtn) {
        clearResultsBtn.addEventListener('click', clearReviewResults);
    }
    
    // 导出下拉菜单事件已在init.js中处理
    
    console.log('review.js 事件监听器已绑定');
});

