function parseWordTextToRecords(text) {
    const records = [];
    try {
        // 保留换行，避免把整段压成一行影响逐行解析
        text = text
            .replace(/\t/g, ' ')
            .replace(/\u00A0/g, ' ')
            .replace(/[ \t\u00A0]+/g, ' ');
        const lines = text.split('\n');
        const specialTableRecords = parseSpecialTableFormat(text);
        console.log('parseSpecialTableFormat 结果:', specialTableRecords);
        if (specialTableRecords.length > 0) { 
            console.log('使用 specialTableRecords，数量:', specialTableRecords.length);
            records.push(...specialTableRecords); 
            return records; 
        }
        const allRecords = extractAllPossibleRecords(lines);
        if (allRecords.length > 0) { records.push(...allRecords); return records; }
        const paragraphs = text.split(/\n\n+/);
        const paragraphRecords = extractParagraphRecords(paragraphs);
        if (paragraphRecords.length > 0) { records.push(...paragraphRecords); return records; }
        let currentRecord = {}; let inActivitySection = false; let activityLines = [];
        lines.forEach((line) => {
            line = line.trim(); if (!line) return;
            if (inActivitySection) { activityLines.push(line); if (line.includes('合计') || line.includes('总分') || line.includes('总计')) { inActivitySection = false; parseActivityLines(activityLines, currentRecord, records); activityLines = []; currentRecord = {}; } return; }
            if (((/姓\s*名/.test(line)) || line.includes('姓名') || line.includes('学生姓名')) && !currentRecord['姓名']) currentRecord['姓名'] = extractField(line, ['姓名', '学生姓名', '姓\\s*名']);
            if ((line.includes('学号') || line.includes('编号')) && !currentRecord['学号']) currentRecord['学号'] = extractField(line, ['学号', '编号'], /\d+/);
            if ((line.includes('所在部门') || line.includes('部门') || line.includes('学院') || line.includes('系别') || line.includes('单位') || line.includes('组织') || line.includes('所属学院') || line.includes('所在学院')) && !currentRecord['所在部门']) currentRecord['所在部门'] = normalizeDepartment(extractField(line, ['所在部门', '部门', '学院', '系别', '单位', '组织', '所属学院', '所在学院']));
            if ((line.includes('实践活动') && line.includes('序号')) || line.includes('所参加的活动') || line.includes('活动名称') || line.includes('活动列表') || (line.includes('序号') && (line.includes('活动') || line.includes('内容')))) { inActivitySection = true; activityLines.push(line); }
            if (Object.keys(currentRecord).length > 0) tryExtractActivityFromLine(line, currentRecord, records);
        });
        if (activityLines.length > 0 && Object.keys(currentRecord).length > 0) parseActivityLines(activityLines, currentRecord, records);
        if (records.length === 0) return parseAlternativeFormat(text);

        // Fallback: 如果记录里缺少姓名/部门，尝试从全文兜底提取一次并回填
        const fallbackName = extractGlobalName(text);
        const fallbackDept = extractGlobalDepartment(text);
        if (fallbackName || fallbackDept) {
            records.forEach(r => {
                if (!r['姓名'] || r['姓名'] === '未知') r['姓名'] = fallbackName || r['姓名'] || '未知';
                if (!r['所在部门'] || r['所在部门'] === '未知') r['所在部门'] = fallbackDept || r['所在部门'] || '未知';
            });
        }
    } catch (error) {
        console.error('解析Word文本出错:', error);
        showNotification('解析Word文档时出错: ' + error.message, 'error');
    }
    return records;
}

function extractField(line, keywords, pattern = null) {
    const valuePatterns = [
        new RegExp(`(${keywords.join('|')})[\\s:：]*([^\\n\\r]+)`),
        new RegExp(`(${keywords.join('|')})[\\s]*=[\\s]*([^\\n\\r]+)`),
        new RegExp(`([^\\n\\r]+?)[\\s]*(${keywords.join('|')})`)
    ];
    for (const valuePattern of valuePatterns) {
        const match = line.match(valuePattern);
        if (match && match.length >= 3) {
            let value = match[2];
            const keywordSet = new Set(keywords);
            if (keywordSet.has(value)) { value = match[1]; }
            value = (value || '').trim();
            if (pattern) { const patternMatch = value.match(pattern); if (patternMatch) value = patternMatch[0]; }
            // 如果是姓名字段，优先截取第一个2-4位中文，避免串入其他字段
            if ([...keywordSet].some(k => /姓名|学生姓名|姓\s*名/.test(k))) {
                const shortName = value.match(/[\u4e00-\u9fa5·]{2,4}/);
                if (shortName) value = shortName[0];
            }
            return value
                .replace(/[\\s\\t\\u00A0]+/g, '')
                .replace(/[^\\u4e00-\\u9fa5a-zA-Z0-9\\-·()（）/\\\\&]/g, '')
                .trim();
        }
    }
    return '';
}

function tryExtractActivityFromLine(line, currentRecord, records) {
    // 跳过包含字段标签的行
    if (line.includes('加分') || line.includes('活动名称') || line.includes('序号') || line.includes('姓名') || line.includes('学号') || line.includes('部门')) {
        return;
    }
    
    // 检查是否包含数字（可能是分数）
    if (/\d/.test(line)) {
        const scoreMatch = line.match(/([\d.]+)$/);
        if (scoreMatch) {
            const score = parseFloat(scoreMatch[1]);
            if (!isNaN(score) && score > 0) {
                let activityName = line.replace(/[\d.]+$/, '').trim();
                console.log('原始活动名称:', activityName);
                
                // 智能移除开头的序号，但保留年份信息
                // 只移除真正的序号（如"1."、"2."等），保留年份（如"2024"）
                // 判断逻辑：如果开头是1-2位数字+点号+空格，则认为是序号
                if (/^\d{1,2}\.\s/.test(activityName)) {
                    // 移除序号（如"1. "、"2. "等）
                    activityName = activityName.replace(/^\d{1,2}\.\s/, '');
                }
                console.log('移除序号后的活动名称:', activityName);
                
                // 过滤掉字段标签和无效内容
                if (activityName.length > 2 && !/加分|活动名称|序号|姓名|学号|部门|学院|系别|单位|组织/.test(activityName)) {
                    const activityRecord = Object.assign({}, currentRecord);
                    activityRecord['所参加的活动及担任角色'] = activityName;
                    activityRecord['加分'] = score;
                    console.log('添加活动记录:', activityRecord);
                    records.push(activityRecord);
                }
            }
        }
    }
}

function extractAllPossibleRecords(lines) {
    const records = [];
    const peopleData = [];
    let currentPerson = {};
    let currentActivities = [];
    lines.forEach(line => {
        line = line.trim();
        if (!line) return;
        if (line.includes('---') || line.includes('====') || line.length > 50 && /[\s]+/.test(line)) {
            if (Object.keys(currentPerson).length > 0 || currentActivities.length > 0) {
                peopleData.push({ person: currentPerson, activities: currentActivities });
                currentPerson = {};
                currentActivities = [];
            }
            return;
        }
        if (line.includes('姓名') && !currentPerson['姓名']) currentPerson['姓名'] = extractField(line, ['姓名']);
        if (line.includes('学号') && !currentPerson['学号']) currentPerson['学号'] = extractField(line, ['学号'], /\d+/);
        if ((line.includes('部门') || line.includes('学院') || line.includes('系别') || line.includes('单位') || line.includes('组织') || line.includes('所属学院') || line.includes('所在学院')) && !currentPerson['所在部门']) currentPerson['所在部门'] = normalizeDepartment(extractField(line, ['部门', '学院', '系别', '单位', '组织', '所属学院', '所在学院']));
        // 跳过包含字段标签的行
        if (line.includes('加分') || line.includes('活动名称') || line.includes('序号') || line.includes('姓名') || line.includes('学号') || line.includes('部门')) {
            return;
        }
        
        const scoreMatch = line.match(/([\d.]+)$/);
        if (scoreMatch) { 
            const score = parseFloat(scoreMatch[1]); 
            if (!isNaN(score) && score > 0) { 
                let activityName = line.replace(/[\d.]+$/, '').trim(); 
                // 智能移除开头的序号，但保留年份信息
                // 只移除真正的序号（如"1."、"2."等），保留年份（如"2024"）
                // 判断逻辑：如果开头是1-2位数字+点号+空格，则认为是序号
                if (/^\d{1,2}\.\s/.test(activityName)) {
                    // 移除序号（如"1. "、"2. "等）
                    activityName = activityName.replace(/^\d{1,2}\.\s/, '');
                }
                // 过滤掉字段标签和无效内容
                if (activityName.length > 2 && !/加分|活动名称|序号|姓名|学号|部门|学院|系别|单位|组织/.test(activityName)) {
                    currentActivities.push({ name: activityName, score: score }); 
                }
            } 
        }
    });
    if (Object.keys(currentPerson).length > 0 || currentActivities.length > 0) {
        peopleData.push({ person: currentPerson, activities: currentActivities });
    }
    peopleData.forEach(data => {
        const person = data.person;
        data.activities.forEach(activity => {
            records.push({
                '姓名': person['姓名'] || '未知',
                '学号': person['学号'] || '未知',
                '所在部门': normalizeDepartment(person['所在部门']) || '未知',
                '所参加的活动及担任角色': activity.name,
                '加分': activity.score
            });
        });
    });
    return records;
}

function extractParagraphRecords(paragraphs) {
    const records = [];
    paragraphs.forEach(paragraph => {
        paragraph = paragraph.trim();
        if (paragraph.length < 20) return;
        const lines = paragraph.split('\n');
        const person = {};
        const activities = [];
        lines.forEach(line => { 
            line = line.trim(); 
            if (!line) return; 
            
            if (((/姓\s*名/.test(line)) || line.includes('姓名') || line.includes('学生姓名')) && !person['姓名']) 
                person['姓名'] = extractField(line, ['姓名', '学生姓名', '姓\\s*名']); 
            if (line.includes('学号') && !person['学号']) 
                person['学号'] = extractField(line, ['学号'], /\d+/); 
            if ((line.includes('部门') || line.includes('学院') || line.includes('系别') || line.includes('单位') || line.includes('组织') || line.includes('所属学院') || line.includes('所在学院')) && !person['所在部门']) 
                person['所在部门'] = normalizeDepartment(extractField(line, ['部门', '学院', '系别', '单位', '组织', '所属学院', '所在学院'])); 
            
            // 跳过包含字段标签的行
            if (line.includes('加分') || line.includes('活动名称') || line.includes('序号') || line.includes('姓名') || line.includes('学号') || line.includes('部门')) {
                return;
            }
            
            const scoreMatch = line.match(/([\d.]+$)/); 
            if (scoreMatch) { 
                const score = parseFloat(scoreMatch[1]); 
                if (!isNaN(score) && score > 0) { 
                    let activityName = line.replace(/[\d.]+$/, '').trim(); 
                    // 智能移除开头的序号，但保留年份信息
                    // 只移除真正的序号（如"1."、"2."等），保留年份（如"2024"）
                    // 判断逻辑：如果开头是1-2位数字+点号+空格，则认为是序号
                    if (/^\d{1,2}\.\s/.test(activityName)) {
                        // 移除序号（如"1. "、"2. "等）
                        activityName = activityName.replace(/^\d{1,2}\.\s/, '');
                    }
                    // 过滤掉字段标签和无效内容
                    if (activityName.length > 2 && !/加分|活动名称|序号|姓名|学号|部门|学院|系别|单位|组织/.test(activityName)) {
                        activities.push({ name: activityName, score: score }); 
                    }
                } 
            } 
        });
        activities.forEach(activity => {
            records.push({
                '姓名': person['姓名'] || '未知',
                '学号': person['学号'] || '未知',
                '所在部门': normalizeDepartment(person['所在部门']) || '未知',
                '所参加的活动及担任角色': activity.name,
                '加分': activity.score
            });
        });
    });
    return records;
}

function parseAlternativeFormat(text) {
    const records = [];
    try {
        const lines = text.split('\n');
        let isHeaderFound = false;
        let headers = [];
        lines.forEach((line) => {
            line = line.trim();
            if (!line) return;
            if (!isHeaderFound && (line.includes('姓名') && line.includes('学号') && line.includes('部门') && line.includes('活动'))) {
                isHeaderFound = true;
                headers = line.split(/\s+/);
                return;
            }
            if (isHeaderFound) {
                const data = line.split(/\t|\s{2,}/);
                if (data.length >= 4) {
                    const record = {};
                    data.forEach((value, i) => {
                        value = value.trim();
                        if (!value) return;
                        if (i === 0 && !record['姓名']) {
                            if (!/^\d+$/.test(value)) record['姓名'] = value;
                        } else if (i === 1 && !record['学号'] && /\d+/.test(value)) {
                            record['学号'] = value;
                        } else if (i === 2 && !record['所在部门']) {
                            record['所在部门'] = normalizeDepartment(value) || '未知';
                        } else if (i === 3 && !record['所参加的活动及担任角色']) {
                            record['所参加的活动及担任角色'] = value;
                        }
                        if (!record['加分']) {
                            const scoreMatch = value.match(/([\d.]+)/);
                            if (scoreMatch) {
                                const score = parseFloat(scoreMatch[1]);
                                if (!isNaN(score)) {
                                    record['加分'] = score;
                                }
                            }
                        }
                    });
                    if (record['姓名'] && record['所参加的活动及担任角色'] && record['加分']) {
                        record['学号'] = record['学号'] || '未知';
                        record['所在部门'] = record['所在部门'] || '未知';
                        records.push(record);
                    }
                }
            }
        });
    } catch (error) {
        console.error('备用解析策略出错:', error);
    }
    return records;
}

function parseActivityLines(lines, currentRecord, records) {
    lines.forEach(line => { 
        line = line.trim(); 
        if (!line || line.includes('合计') || line.includes('序号') || line.includes('总分') || line.includes('总计')) return; 
        
        // 跳过包含字段标签的行
        if (line.includes('加分') || line.includes('活动名称') || line.includes('姓名') || line.includes('学号') || line.includes('部门')) {
            return;
        }
        
        if (/^\d+[\s\t]+/.test(line)) { 
            let activityMatch = line.match(/^\d+[\s\t]+(.+?)[\s\t]+(\d+(?:\.\d+)?)$/); 
            if (activityMatch && activityMatch.length >= 3) { 
                const activityName = activityMatch[1].trim();
                // 过滤掉字段标签和无效内容
                if (activityName.length > 2 && !/加分|活动名称|序号|姓名|学号|部门|学院|系别|单位|组织/.test(activityName)) {
                    const activityRecord = Object.assign({}, currentRecord); 
                    activityRecord['所参加的活动及担任角色'] = activityName; 
                    activityRecord['加分'] = parseFloat(activityMatch[2]); 
                    records.push(activityRecord); 
                }
                return; 
            } 
                        // 智能移除开头的序号，但保留年份信息
            // 只移除真正的序号（如"1."、"2."等），保留年份（如"2024"）
            // 判断逻辑：如果开头是1-2位数字+点号+空格，则认为是序号
            let activityName = line;
            if (/^\d{1,2}\.\s/.test(line)) {
                // 移除序号（如"1. "、"2. "等）
                activityName = line.replace(/^\d{1,2}\.\s/, '');
            }
            console.log('parseActivityLines - 移除序号后的活动名称:', activityName);
            
            const scoreMatch = activityName.match(/([\d.]+)$/); 
            if (scoreMatch) { 
                activityName = activityName.replace(/[\d.]+$/, '').trim();
                console.log('parseActivityLines - 最终活动名称:', activityName);
                
                // 过滤掉字段标签和无效内容
                if (activityName.length > 2 && !/加分|活动名称|序号|姓名|学号|部门|学院|系别|单位|组织/.test(activityName)) {
                    const activityRecord = Object.assign({}, currentRecord); 
                    activityRecord['加分'] = parseFloat(scoreMatch[1]); 
                    activityRecord['所参加的活动及担任角色'] = activityName; 
                    console.log('parseActivityLines - 添加活动记录:', activityRecord);
                    records.push(activityRecord); 
                }
            } 
        } else { 
            const scoreMatch = line.match(/(.+?)[\s\t]+(\d+(?:\.\d+)?)$/); 
            if (scoreMatch && scoreMatch.length >= 3) { 
                const activityName = scoreMatch[1].trim();
                // 过滤掉字段标签和无效内容
                if (activityName.length > 2 && !/加分|活动名称|序号|姓名|学号|部门|学院|系别|单位|组织/.test(activityName)) {
                    const activityRecord = Object.assign({}, currentRecord); 
                    activityRecord['所参加的活动及担任角色'] = activityName; 
                    activityRecord['加分'] = parseFloat(scoreMatch[2]); 
                    records.push(activityRecord); 
                }
            } 
        } 
    });
}

function parseSpecialTableFormat(text) {
    const records = [];
    try {
        let name = '';
        const directNameMatch = text.match(/(?:姓名|姓\s*名)\s+([\u4e00-\u9fa5·]{2,4})/);
        if (directNameMatch && directNameMatch[1]) name = directNameMatch[1];
        if (!name) {
            const namePatterns = [/姓名\s*[：:]\s*([\u4e00-\u9fa5·]{2,4})/, /姓\s*名\s*[：:]\s*([\u4e00-\u9fa5·]{2,4})/, /姓名\s+([\u4e00-\u9fa5·]{2,4})/, /姓名([\u4e00-\u9fa5·]{2,4})/, /姓名\s*[：:]\s*([^\s\n\r]+)/, /姓\s*名\s*[：:]\s*([^\s\n\r]+)/, /姓名([^\s\n\r]+)/];
            for (const pattern of namePatterns) {
                const match = text.match(pattern);
                if (match && match[1]) {
                    const extractedName = match[1].trim();
                    if (extractedName && extractedName.length >= 2 && extractedName.length <= 4) {
                        name = extractedName;
                        break;
                    }
                }
            }
        }
        if (!name) { name = extractGlobalName(text); }
        let studentId = '未知';
        const sidIndex = text.search(/学号/);
        if (sidIndex !== -1) {
            const ctx = text.slice(Math.max(0, sidIndex), sidIndex + 120);
            const fieldMatch = ctx.match(/学号[^\d\n\r]*([0-9]{6,20})/);
            if (fieldMatch && fieldMatch[1]) studentId = fieldMatch[1];
        }
        if (studentId === '未知') {
            const candidates = text.match(/\b\d{8,12}\b/g) || [];
            if (candidates.length === 1) studentId = candidates[0];
        }
        let department = '未知';
        const deptFieldIdx = text.search(/所在部门|部门\s*[：:]/);
        if (deptFieldIdx !== -1) {
            const context = text.slice(Math.max(0, deptFieldIdx - 60), deptFieldIdx + 240);
            const colonMatch = context.match(/(?:所在部门|部门)[^\n\r：:]*[：:]\s*([\u4e00-\u9fa5·\s、，,／\/\\-]{2,30})/);
            if (colonMatch && colonMatch[1]) {
                let raw = colonMatch[1].trim();
                raw = raw.replace(/(职务|干事|部长|副部长|秘书长|主席|副主席)/g, '');
                const orgMatch = raw.match(/[\u4e00-\u9fa5·]{2,15}(?:部|学院|系|处|中心)/);
                if (orgMatch) department = orgMatch[0];
            } else {
                const noColon = context.match(/所在部门[^\n\r]{0,10}[、，,\s]+[^\n\r]{0,12}?([\u4e00-\u9fa5·]{2,15}(?:部|学院|系|处|中心))/);
                if (noColon && noColon[1]) department = noColon[1];
            }
        }
        if (!department || department === '未知') {
            department = extractGlobalDepartment(text);
        }
        const activities = []; 
        // 使用更简单有效的正则表达式来匹配活动名称和分数
        // 匹配格式：活动名称 + 空格/制表符 + 分数
        // 活动名称可以包含数字（如年份），但会过滤掉以序号开头的行
        const activityPattern = /([^0-9\n\r]*\d{4}[^0-9\n\r]*?|[^0-9\n\r]+?)\s+(\d+(?:\.\d+)?)/g;
        console.log('parseSpecialTableFormat - 开始解析活动，文本长度:', text.length);
        console.log('parseSpecialTableFormat - 文本预览:', text.substring(0, 200)); 
        let match; 
        while ((match = activityPattern.exec(text)) !== null) { 
            let activityName = match[1].trim(); 
            const score = parseFloat(match[2]); 
            console.log('parseSpecialTableFormat - 原始活动名称:', activityName);
            
            // 智能移除开头的序号，但保留年份信息
            // 只移除真正的序号（如"1."、"2."等），保留年份（如"2024"）
            // 判断逻辑：如果开头是1-2位数字+点号+空格，则认为是序号
            if (/^\d{1,2}\.\s/.test(activityName)) {
                // 移除序号（如"1. "、"2. "等）
                activityName = activityName.replace(/^\d{1,2}\.\s/, '');
            } 
            console.log('parseSpecialTableFormat - 移除序号后的活动名称:', activityName);
            
            activityName = activityName.replace(/[。！？，、\s]+$/, ''); 
            // 过滤掉字段标签和无效内容
            if (activityName.length > 1 && !isNaN(score) && score > 0 && score <= 10 && 
                !/加分|活动名称|序号|姓名|学号|部门|学院|系别|单位|组织/.test(activityName)) {
                console.log('parseSpecialTableFormat - 最终活动名称:', activityName);
                activities.push({ name: activityName, score }); 
            }
        }
        if (activities.length > 0) {
            activities.forEach(activity => {
                records.push({
                    '姓名': name || '未知',
                    '学号': studentId,
                    '所在部门': normalizeDepartment(department) || '未知',
                    '所参加的活动及担任角色': activity.name,
                    '加分': activity.score
                });
            });
        }
    } catch (error) {
        console.error('特殊表格格式解析出错:', error);
    }
    return records;
}

// 全局兜底：从全文中尽量提取姓名（2-4个连续中文）
function extractGlobalName(text) {
    // 优先匹配带关键词（容忍换行/空白/标点间隔），并允许值在下一行
    const keyContext = text.match(/(?:姓\s*名|姓名|学生姓名)[^\n\r：:]{0,10}[：: ]?\s*([\s\S]{0,20})/);
    if (keyContext && keyContext[1]) {
        const ctx = keyContext[1];
        const nameMatch = ctx.match(/[\u4e00-\u9fa5·]\s*[\u4e00-\u9fa5·]{1,3}/);
        if (nameMatch) return nameMatch[0].replace(/\s+/g, '').trim();
    }
    // 行级匹配：关键字下一行是名字
    const lines = text.split(/\r?\n/);
    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        if (/姓\s*名|姓名|学生姓名/.test(line)) {
            // 先尝试本行
            const inline = line.match(/(?:姓\s*名|姓名|学生姓名)[^\n\r：:]{0,10}[：: ]?\s*([\u4e00-\u9fa5·]\s*[\u4e00-\u9fa5·]{1,3})/);
            if (inline && inline[1]) return inline[1].replace(/\s+/g, '').trim();
            // 再看下一行的前若干字符
            const next = (lines[i + 1] || '').trim();
            const nextName = next.match(/^[\u4e00-\u9fa5·]\s*[\u4e00-\u9fa5·]{1,3}/);
            if (nextName && nextName[0]) return nextName[0].replace(/\s+/g, '').trim();
        }
    }
    // 兼容“姓 名”中间有空格/制表
    const spaced = text.match(/姓\s*名\s*[：: ]\s*([\u4e00-\u9fa5·]\s*[\u4e00-\u9fa5·]{1,3})/);
    if (spaced && spaced[1]) return spaced[1].replace(/\s+/g, '').trim();
    // 退化：挑选最可能的人名（过滤常见词）
    const candidates = (text.match(/[\u4e00-\u9fa5·]{2,4}/g) || []).filter(t => !/姓名|学号|部门|学院|性别|民族|政治|出生|班级|专业|活动|角色|分|合计|总分|总计/.test(t));
    return candidates.length ? candidates[0].replace(/\s+/g, '') : '';
}

// 全局兜底：从全文中尽量提取部门（包含“部/学院/系/处/中心”等后缀）
function extractGlobalDepartment(text) {
    // 带关键词的优先（允许值在下一行）
    const keyCtx = text.match(/(?:所在\s*部门|部门|单位|组织|学院|所属学院|所在学院)[^\n\r：:]{0,10}[：: ]?\s*([\s\S]{0,40})/);
    if (keyCtx && keyCtx[1]) {
        const ctx = keyCtx[1];
        const deptMatch = ctx.match(/[\u4e00-\u9fa5·（）()\-/&]{2,30}?(?:部|学院|系|处|中心|科|组|办|队)/);
        if (deptMatch) return deptMatch[0].replace(/\s+/g, '').trim();
    }
    // 行级匹配：关键字下一行
    const lines = text.split(/\r?\n/);
    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        if (/所在\s*部门|部门|单位|组织|学院|所属学院|所在学院/.test(line)) {
            const inline = line.match(/(?:所在\s*部门|部门|单位|组织|学院|所属学院|所在学院)[^\n\r：:]{0,10}[：: ]?\s*([\u4e00-\u9fa5·（）()\-/&]{2,30}?(?:部|学院|系|处|中心|科|组|办|队))/);
            if (inline && inline[1]) return inline[1].replace(/\s+/g, '').trim();
            const next = (lines[i + 1] || '').trim();
            const nextDept = next.match(/^[\u4e00-\u9fa5·（）()\-/&]{2,30}?(?:部|学院|系|处|中心|科|组|办|队)/);
            if (nextDept && nextDept[0]) return nextDept[0].replace(/\s+/g, '').trim();
        }
    }
    // 无冒号格式：在“所在部门”邻近区域提取
    const near = text.match(/所在部门[\s\S]{0,80}?([\u4e00-\u9fa5·（）()\-/&]{2,30}(?:部|学院|系|处|中心|科|组|办|队))/);
    if (near && near[1]) return near[1].replace(/\s+/g, '').trim();
    // 退化：全局挑选一个组织后缀词
    const any = text.match(/[\u4e00-\u9fa5·（）()\-/&]{2,30}(?:部|学院|系|处|中心|科|组|办|队)/);
    return any ? normalizeDepartment(any[0]) : '';
}

// 规范化部门值：去掉职务等后缀，只保留组织名
function normalizeDepartment(value) {
    if (!value) return value;
    let v = String(value).replace(/\s+/g, '');
    // 先移除明显的非组织词，若只剩这些则视为未填写
    v = v.replace(/所在部门|单位|组织|学院|所属学院|所在学院|部门|：|:/g, '');
    // 过滤职位/身份等：包含这些直接判定为空
    if (/(干部|干事|职务|岗位|部长|副部长|秘书长|主任|副主任|主席|副主席|委员|成员)$/.test(v)) {
        // 若是“图书馆学生干部”等，判定为未填写
        return '';
    }
    // 避免把“干部”的“部”当作组织后缀
    v = v.replace(/干部/g, '');
    v = v.replace(/干事/g, '');
    // 若含有空格/分隔符，优先截取组织名在前的部分
    const m = v.match(/([\u4e00-\u9fa5·（）()\-/&]{2,30}?(?:部|学院|系|处|中心|科|组|办|队))/);
    // 若找不到带组织后缀的词，则认定未填写
    return m ? m[1] : '';
}

// 解析单个文件（doc/docx/xlsx/xls/csv）
window.processFile = function(file) {
    return new Promise((resolve) => {
        const ext = file.name.split('.').pop().toLowerCase();
        if (['doc', 'docx'].includes(ext)) {
            const reader = new FileReader();
            reader.onload = function(e) {
                const arrayBuffer = e.target.result;
                mammoth.extractRawText({ arrayBuffer })
                    .then(async function(result) {
                        let text = result.value || '';
                        // 额外：对嵌入图片做OCR，补充到文本中（可选）
                        try {
                            const zip = new JSZip();
                            await zip.loadAsync(arrayBuffer);
                            const imageFolderNames = Object.keys(zip.files).filter(p => p.startsWith('word/media/'));
                            if (imageFolderNames.length > 0 && window.Tesseract) {
                                for (const imgPath of imageFolderNames) {
                                    const file = zip.files[imgPath];
                                    if (!file) continue;
                                    const blob = await file.async('blob');
                                    const dataUrl = await new Promise(res => {
                                        const reader = new FileReader(); reader.onload = () => res(reader.result); reader.readAsDataURL(blob);
                                    });
                                    try {
                                        const ocr = await Tesseract.recognize(dataUrl, 'chi_sim+eng', { logger: () => {} });
                                        const ocrText = (ocr && ocr.data && ocr.data.text) ? ocr.data.text : '';
                                        if (ocrText && ocrText.trim().length > 0) {
                                            text += `\n${ocrText}`;
                                        }
                                    } catch (e) { /* 忽略单张图片OCR错误 */ }
                                }
                            }
                        } catch (e) { /* 如果没有JSZip或解析失败则忽略OCR流程，不影响文本解析 */ }
                        const records = parseWordTextToRecords(text);
                        console.log('Word文档解析结果:', records);
                        records.forEach(record => {
                            console.log('处理记录:', record);
                            const reviewResult = reviewRecord(record);
                            reviewResults.push(reviewResult);
                        });
                        resolve();
                    })
                    .catch(function(error) {
                        console.error('解析Word文档出错:', error);
                        showNotification('解析Word文档时出错: ' + error.message, 'error');
                        resolve();
                    });
            };
            reader.onerror = function(error) {
                console.error('文件读取错误:', error);
                showNotification('读取文件时出错', 'error');
                resolve();
            };
            reader.readAsArrayBuffer(file);
        } else if (['xlsx', 'xls', 'csv'].includes(ext)) {
            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const sheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[sheetName];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet);
                    jsonData.forEach(row => {
                        const result = reviewRecord(row);
                        reviewResults.push(result);
                    });
                } catch (error) {
                    console.error('解析Excel/CSV文件出错:', error);
                    showNotification('解析Excel/CSV文件时出错: ' + error.message, 'error');
                }
                resolve();
            };
            reader.onerror = function(error) {
                console.error('文件读取错误:', error);
                showNotification('读取文件时出错', 'error');
                resolve();
            };
            reader.readAsArrayBuffer(file);
        } else {
            console.log('不支持的文件类型:', ext);
            resolve();
        }
    });
}

