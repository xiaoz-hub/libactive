function parseWordTextToRecords(text) {
    const records = [];
    try {
        // 保留换行，避免把整段压成一行影响逐行解析
        text = text
            .replace(/\t/g, ' ')
            .replace(/\u00A0/g, ' ')
            .replace(/[ \t\u00A0]+/g, ' ');
        const lines = text.split('\n');
        
        // 首先尝试特殊表格格式
        const specialTableRecords = parseSpecialTableFormat(text);
        console.log('parseSpecialTableFormat 结果:', specialTableRecords);
        if (specialTableRecords.length > 0) { 
            console.log('使用 specialTableRecords，数量:', specialTableRecords.length);
            records.push(...specialTableRecords); 
            return records; 
        }
        
        // 尝试提取所有可能的记录
        const allRecords = extractAllPossibleRecords(lines);
        if (allRecords.length > 0) { 
            console.log('使用 extractAllPossibleRecords，数量:', allRecords.length);
            records.push(...allRecords); 
            return records; 
        }
        
        // 尝试段落格式
        const paragraphs = text.split(/\n\n+/);
        const paragraphRecords = extractParagraphRecords(paragraphs);
        if (paragraphRecords.length > 0) { 
            console.log('使用 extractParagraphRecords，数量:', paragraphRecords.length);
            records.push(...paragraphRecords); 
            return records; 
        }
        
        // 最后尝试逐行解析
        let currentRecord = {}; 
        let inActivitySection = false; 
        let activityLines = [];
        let extractedCount = 0;
        
        lines.forEach((line, lineIndex) => {
            line = line.trim(); 
            if (!line) return;
            
            // 如果在活动区域内
            if (inActivitySection) { 
                activityLines.push(line); 
                // 检查是否到达活动区域结束
                if (line.includes('总分') || line.includes('总计')) { 
                    inActivitySection = false; 
                    console.log('活动区域结束，处理活动行数:', activityLines.length);
                    parseActivityLines(activityLines, currentRecord, records); 
                    activityLines = []; 
                    currentRecord = {}; 
                } 
                
                // 跳过合计行，因为合计只是总分的计算，不需要提取
                if (line.includes('合计')) {
                    console.log('跳过合计行（不需要提取）:', line);
                    return;
                }
                return; // 在活动区域内不进行即时提取
            }
            
            // 提取基本信息
            if (((/姓\s*名/.test(line)) || line.includes('姓名') || line.includes('学生姓名')) && !currentRecord['姓名']) {
                currentRecord['姓名'] = extractField(line, ['姓名', '学生姓名', '姓\\s*名']);
            }
            if ((line.includes('学号') || line.includes('编号')) && !currentRecord['学号']) {
                const extractedStudentId = extractField(line, ['学号', '编号'], /\d+/);
                if (extractedStudentId) {
                    currentRecord['学号'] = extractedStudentId;
                } else {
                    // 如果标准提取失败，尝试更宽松的匹配
                    const studentIdMatch = line.match(/学号[^\d]*(\d{6,20})/);
                    if (studentIdMatch && studentIdMatch[1]) {
                        currentRecord['学号'] = studentIdMatch[1];
                    }
                }
            }
            if ((line.includes('所在部门') || line.includes('部门') || line.includes('学院') || line.includes('系别') || line.includes('单位') || line.includes('组织') || line.includes('所属学院') || line.includes('所在学院')) && !currentRecord['所在部门']) {
                currentRecord['所在部门'] = normalizeDepartment(extractField(line, ['所在部门', '部门', '学院', '系别', '单位', '组织', '所属学院', '所在学院']));
            }
            
            // 检测活动区域开始
            if ((line.includes('实践活动') && line.includes('序号')) || 
                line.includes('所参加的活动') || 
                line.includes('活动名称') || 
                line.includes('活动列表') || 
                (line.includes('序号') && (line.includes('活动') || line.includes('内容')))) { 
                inActivitySection = true; 
                activityLines.push(line); 
                console.log('检测到活动区域开始，行:', lineIndex + 1);
                return; 
            }
            
            // 如果不在活动区域内，尝试即时提取活动
            if (Object.keys(currentRecord).length > 0) {
                const beforeCount = records.length;
                tryExtractActivityFromLine(line, currentRecord, records);
                if (records.length > beforeCount) {
                    extractedCount++;
                    console.log('即时提取活动成功，行:', lineIndex + 1);
                }
            }
        });
        
        // 处理剩余的活动行
        if (activityLines.length > 0 && Object.keys(currentRecord).length > 0) {
            console.log('处理剩余活动行，数量:', activityLines.length);
            parseActivityLines(activityLines, currentRecord, records);
        }
        
        // 如果所有方法都失败，尝试备用格式
        if (records.length === 0) {
            console.log('尝试备用解析格式');
            const alternativeRecords = parseAlternativeFormat(text);
            if (alternativeRecords.length > 0) {
                records.push(...alternativeRecords);
            }
        }

        // Fallback: 如果记录里缺少姓名/部门，尝试从全文兜底提取一次并回填
        const fallbackName = extractGlobalName(text);
        const fallbackDept = extractGlobalDepartment(text);
        if (fallbackName || fallbackDept) {
            records.forEach(r => {
                if (!r['姓名'] || r['姓名'] === '未知') r['姓名'] = fallbackName || r['姓名'] || '未知';
                if (!r['所在部门'] || r['所在部门'] === '未知') r['所在部门'] = fallbackDept || r['所在部门'] || '未知';
            });
        }
        
        // 去重处理：移除重复的活动记录
        const uniqueRecords = [];
        const seenActivities = new Set();
        
        records.forEach(record => {
            const key = `${record['姓名']}-${record['所参加的活动及担任角色']}-${record['加分']}`;
            if (!seenActivities.has(key)) {
                seenActivities.add(key);
                uniqueRecords.push(record);
            } else {
                console.log('移除重复记录:', record);
            }
        });
        
        console.log('去重前记录数量:', records.length);
        console.log('去重后记录数量:', uniqueRecords.length);
        console.log('最终提取结果，总数量:', uniqueRecords.length);
        
    } catch (error) {
        console.error('解析Word文本出错:', error);
        showNotification('解析Word文档时出错: ' + error.message, 'error');
    }
    return uniqueRecords;
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
            if (pattern) { 
                const patternMatch = value.match(pattern); 
                if (patternMatch) {
                    value = patternMatch[0];
                    // 对于学号，确保提取的是完整的学号
                    if (keywords.some(k => k.includes('学号'))) {
                        // 如果学号被截断了，尝试提取更长的学号
                        const longerMatch = value.match(/\d{6,20}/);
                        if (longerMatch) {
                            value = longerMatch[0];
                        }
                    }
                }
            }
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
    // 跳过包含字段标签的行，但允许活动名称中包含这些词
    if (line.trim().match(/^(加分|活动名称|序号|姓名|学号|部门|学院|系别|单位|组织)[：:]/)) {
        return;
    }
    
    // 跳过明显的活动区域标记行
    if (line.includes('所参加的活动') || line.includes('活动列表') || line.includes('活动记录')) {
        return;
    }
    
    // 跳过合计行，因为合计只是总分的计算，不需要提取
    if (line.includes('合计') || line.includes('总计') || line.includes('总分')) {
        console.log('跳过合计行（不需要提取）:', line);
        return;
    }
    
    // 改进的分数匹配正则表达式，支持更多格式
    const scorePatterns = [
        /([\d.]+)\s*$/,           // 行末分数
        /[\s\t]+([\d.]+)\s*$/,    // 制表符或空格分隔的分数
        /[：:]\s*([\d.]+)\s*$/,   // 冒号分隔的分数
        /[（(]\s*([\d.]+)\s*[）)]\s*$/, // 括号中的分数
    ];
    
    let score = null;
    let activityName = line;
    
    // 尝试多种分数匹配模式
    for (const pattern of scorePatterns) {
        const scoreMatch = line.match(pattern);
        if (scoreMatch) {
            score = parseFloat(scoreMatch[1]);
            if (!isNaN(score) && score > 0) {
                // 移除分数部分，获取活动名称
                activityName = line.replace(pattern, '').trim();
                break;
            }
        }
    }
    
    if (score && !isNaN(score) && score > 0) {
        console.log('原始活动名称:', activityName);
        
        // 改进的序号处理逻辑
        // 支持更多序号格式：1. 1、 1) 1） (1) （1）等
        const numberPatterns = [
            /^\d{1,2}\.\s/,      // 1. 
            /^\d{1,2}、\s/,      // 1、
            /^\d{1,2}\)\s/,      // 1)
            /^\d{1,2}）\s/,      // 1）
            /^\(\d{1,2}\)\s/,    // (1)
            /^（\d{1,2}）\s/,    // （1）
            /^\d{1,2}\s/,        // 1 (单个数字后跟空格)
        ];
        
        for (const pattern of numberPatterns) {
            if (pattern.test(activityName)) {
                activityName = activityName.replace(pattern, '');
                break;
            }
        }
        
        console.log('移除序号后的活动名称:', activityName);
        
        // 清理活动名称
        activityName = activityName
            .replace(/^[^\u4e00-\u9fa5a-zA-Z0-9]*/, '') // 移除开头的非文字数字字符
            .replace(/[^\u4e00-\u9fa5a-zA-Z0-9\s\-_()（）【】\[\]""''，,。.！!？?]*$/, '') // 移除结尾的无效字符
            .trim();
        
        // 放宽过滤条件，只过滤明显的表头行
        const isHeaderRow = /^(序号|活动名称|加分|姓名|学号|部门|学院|系别|单位|组织)\s*[：:]/;
        const isTooShort = activityName.length < 2;
        const isOnlyNumbers = /^\d+$/.test(activityName);
        const isOnlyPunctuation = /^[^\u4e00-\u9fa5a-zA-Z0-9]+$/.test(activityName);
        const isTotalRow = /^(合计|总计|总分|小计|总计得分|同意加分)/.test(activityName);
        
        // 检查是否已经存在相同的活动记录
        const existingRecord = records.find(r => 
            r['所参加的活动及担任角色'] === activityName && 
            r['加分'] === score &&
            r['姓名'] === currentRecord['姓名']
        );
        
        if (!isHeaderRow.test(activityName) && !isTooShort && !isOnlyNumbers && !isOnlyPunctuation && !isTotalRow && !existingRecord) {
            const activityRecord = Object.assign({}, currentRecord);
            activityRecord['所参加的活动及担任角色'] = activityName;
            activityRecord['加分'] = score;
            console.log('添加活动记录:', activityRecord);
            records.push(activityRecord);
        } else if (existingRecord) {
            console.log('跳过重复的活动记录:', activityName);
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
        if (line.includes('学号') && !currentPerson['学号']) {
            const extractedStudentId = extractField(line, ['学号'], /\d+/);
            if (extractedStudentId) {
                currentPerson['学号'] = extractedStudentId;
            } else {
                // 如果标准提取失败，尝试更宽松的匹配
                const studentIdMatch = line.match(/学号[^\d]*(\d{6,20})/);
                if (studentIdMatch && studentIdMatch[1]) {
                    currentPerson['学号'] = studentIdMatch[1];
                }
            }
        }
        if ((line.includes('部门') || line.includes('学院') || line.includes('系别') || line.includes('单位') || line.includes('组织') || line.includes('所属学院') || line.includes('所在学院')) && !currentPerson['所在部门']) currentPerson['所在部门'] = normalizeDepartment(extractField(line, ['部门', '学院', '系别', '单位', '组织', '所属学院', '所在学院']));
        // 跳过包含字段标签的行
        if (line.includes('加分') || line.includes('活动名称') || line.includes('序号') || line.includes('姓名') || line.includes('学号') || line.includes('部门')) {
            return;
        }
        
        // 跳过合计行，因为合计只是总分的计算，不需要提取
        if (line.includes('合计') || line.includes('总计') || line.includes('总分')) {
            return;
        }
        
        // 改进的分数匹配正则表达式，支持更多格式
        const scorePatterns = [
            /([\d.]+)\s*$/,           // 行末分数
            /[\s\t]+([\d.]+)\s*$/,    // 制表符或空格分隔的分数
            /[：:]\s*([\d.]+)\s*$/,   // 冒号分隔的分数
            /[（(]\s*([\d.]+)\s*[）)]\s*$/, // 括号中的分数
        ];
        
        let score = null;
        let activityName = line;
        
        // 尝试多种分数匹配模式
        for (const pattern of scorePatterns) {
            const scoreMatch = line.match(pattern);
            if (scoreMatch) {
                score = parseFloat(scoreMatch[1]);
                if (!isNaN(score) && score > 0) {
                    // 移除分数部分，获取活动名称
                    activityName = line.replace(pattern, '').trim();
                    break;
                }
            }
        }
        
        if (score && !isNaN(score) && score > 0) { 
            // 改进的序号处理逻辑
            // 支持更多序号格式：1. 1、 1) 1） (1) （1）等
            const numberPatterns = [
                /^\d{1,2}\.\s/,      // 1. 
                /^\d{1,2}、\s/,      // 1、
                /^\d{1,2}\)\s/,      // 1)
                /^\d{1,2}）\s/,      // 1）
                /^\(\d{1,2}\)\s/,    // (1)
                /^（\d{1,2}）\s/,    // （1）
                /^\d{1,2}\s/,        // 1 (单个数字后跟空格)
            ];
            
            for (const pattern of numberPatterns) {
                if (pattern.test(activityName)) {
                    activityName = activityName.replace(pattern, '');
                    break;
                }
            }
            
            // 清理活动名称
            activityName = activityName
                .replace(/^[^\u4e00-\u9fa5a-zA-Z0-9]*/, '') // 移除开头的非文字数字字符
                .replace(/[^\u4e00-\u9fa5a-zA-Z0-9\s\-_()（）【】\[\]""''，,。.！!？?]*$/, '') // 移除结尾的无效字符
                .trim();
            
            // 放宽过滤条件，只过滤明显的表头行
            const isHeaderRow = /^(序号|活动名称|加分|姓名|学号|部门|学院|系别|单位|组织)\s*[：:]/;
            const isTooShort = activityName.length < 2;
            const isOnlyNumbers = /^\d+$/.test(activityName);
            const isOnlyPunctuation = /^[^\u4e00-\u9fa5a-zA-Z0-9]+$/.test(activityName);
            const isTotalRow = /^(合计|总计|总分|小计|总计得分|同意加分)/.test(activityName);
            
            if (!isHeaderRow.test(activityName) && !isTooShort && !isOnlyNumbers && !isOnlyPunctuation && !isTotalRow) {
                currentActivities.push({ name: activityName, score: score }); 
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
            if (line.includes('学号') && !person['学号']) {
                const extractedStudentId = extractField(line, ['学号'], /\d+/);
                if (extractedStudentId) {
                    person['学号'] = extractedStudentId;
                } else {
                    // 如果标准提取失败，尝试更宽松的匹配
                    const studentIdMatch = line.match(/学号[^\d]*(\d{6,20})/);
                    if (studentIdMatch && studentIdMatch[1]) {
                        person['学号'] = studentIdMatch[1];
                    }
                }
            } 
            if ((line.includes('部门') || line.includes('学院') || line.includes('系别') || line.includes('单位') || line.includes('组织') || line.includes('所属学院') || line.includes('所在学院')) && !person['所在部门']) 
                person['所在部门'] = normalizeDepartment(extractField(line, ['部门', '学院', '系别', '单位', '组织', '所属学院', '所在学院'])); 
            
            // 跳过包含字段标签的行
            if (line.includes('加分') || line.includes('活动名称') || line.includes('序号') || line.includes('姓名') || line.includes('学号') || line.includes('部门')) {
                return;
            }
            
            // 跳过合计行，因为合计只是总分的计算，不需要提取
            if (line.includes('合计') || line.includes('总计') || line.includes('总分')) {
                return;
            }
            
            // 改进的分数匹配正则表达式，支持更多格式
            const scorePatterns = [
                /([\d.]+)\s*$/,           // 行末分数
                /[\s\t]+([\d.]+)\s*$/,    // 制表符或空格分隔的分数
                /[：:]\s*([\d.]+)\s*$/,   // 冒号分隔的分数
                /[（(]\s*([\d.]+)\s*[）)]\s*$/, // 括号中的分数
            ];
            
            let score = null;
            let activityName = line;
            
            // 尝试多种分数匹配模式
            for (const pattern of scorePatterns) {
                const scoreMatch = line.match(pattern);
                if (scoreMatch) {
                    score = parseFloat(scoreMatch[1]);
                    if (!isNaN(score) && score > 0) {
                        // 移除分数部分，获取活动名称
                        activityName = line.replace(pattern, '').trim();
                        break;
                    }
                }
            }
            
            if (score && !isNaN(score) && score > 0) { 
                // 改进的序号处理逻辑
                // 支持更多序号格式：1. 1、 1) 1） (1) （1）等
                const numberPatterns = [
                    /^\d{1,2}\.\s/,      // 1. 
                    /^\d{1,2}、\s/,      // 1、
                    /^\d{1,2}\)\s/,      // 1)
                    /^\d{1,2}）\s/,      // 1）
                    /^\(\d{1,2}\)\s/,    // (1)
                    /^（\d{1,2}）\s/,    // （1）
                    /^\d{1,2}\s/,        // 1 (单个数字后跟空格)
                ];
                
                for (const pattern of numberPatterns) {
                    if (pattern.test(activityName)) {
                        activityName = activityName.replace(pattern, '');
                        break;
                    }
                }
                
                // 清理活动名称
                activityName = activityName
                    .replace(/^[^\u4e00-\u9fa5a-zA-Z0-9]*/, '') // 移除开头的非文字数字字符
                    .replace(/[^\u4e00-\u9fa5a-zA-Z0-9\s\-_()（）【】\[\]""''，,。.！!？?]*$/, '') // 移除结尾的无效字符
                    .trim();
                
                // 放宽过滤条件，只过滤明显的表头行
                const isHeaderRow = /^(序号|活动名称|加分|姓名|学号|部门|学院|系别|单位|组织)\s*[：:]/;
                const isTooShort = activityName.length < 2;
                const isOnlyNumbers = /^\d+$/.test(activityName);
                const isOnlyPunctuation = /^[^\u4e00-\u9fa5a-zA-Z0-9]+$/.test(activityName);
                const isTotalRow = /^(合计|总计|总分|小计|总计得分|同意加分)/.test(activityName);
                
                if (!isHeaderRow.test(activityName) && !isTooShort && !isOnlyNumbers && !isOnlyPunctuation && !isTotalRow) {
                    activities.push({ name: activityName, score: score }); 
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
                        // 检查活动名称是否为合计行
                        const activityName = record['所参加的活动及担任角色'];
                        const isTotalRow = /^(合计|总计|总分|小计|总计得分|同意加分)/.test(activityName);
                        
                        if (!isTotalRow) {
                            record['学号'] = record['学号'] || '未知';
                            record['所在部门'] = record['所在部门'] || '未知';
                            records.push(record);
                        } else {
                            console.log('跳过合计行:', activityName);
                        }
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
    console.log('parseActivityLines 开始处理，行数:', lines.length);
    console.log('当前记录信息:', currentRecord);
    
    let processedCount = 0;
    let skippedCount = 0;
    
    lines.forEach((line, index) => { 
        line = line.trim(); 
        if (!line) {
            console.log(`行 ${index + 1}: 空行，跳过`);
            return;
        }
        
        // 跳过明显的结束标记
        if (line.includes('序号') || line.includes('总分') || line.includes('总计')) {
            console.log(`行 ${index + 1}: 结束标记 "${line}"，跳过`);
            return;
        }
        
        // 跳过合计行，因为合计只是总分的计算，不需要提取
        if (line.includes('合计')) {
            console.log(`行 ${index + 1}: 合计行 "${line}"，跳过（不需要提取）`);
            skippedCount++;
            return;
        }
        
        // 跳过明显的表头行，但允许活动名称中包含这些词
        if (line.trim().match(/^(加分|活动名称|序号|姓名|学号|部门|学院|系别|单位|组织)[：:]/)) {
            console.log(`行 ${index + 1}: 表头行 "${line}"，跳过`);
            skippedCount++;
            return;
        }
        
        console.log(`行 ${index + 1}: 处理 "${line}"`);
        
        // 处理带序号的表格格式
        if (/^\d+[\s\t]+/.test(line)) { 
            console.log(`行 ${index + 1}: 检测到序号格式`);
            let activityMatch = line.match(/^\d+[\s\t]+(.+?)[\s\t]+(\d+(?:\.\d+)?)$/); 
            if (activityMatch && activityMatch.length >= 3) { 
                const activityName = activityMatch[1].trim();
                const score = parseFloat(activityMatch[2]);
                
                console.log(`行 ${index + 1}: 序号格式匹配成功，活动名称: "${activityName}", 分数: ${score}`);
                
                // 清理活动名称
                const cleanedName = activityName
                    .replace(/^[^\u4e00-\u9fa5a-zA-Z0-9]*/, '')
                    .replace(/[^\u4e00-\u9fa5a-zA-Z0-9\s\-_()（）【】\[\]""''，,。.！!？?]*$/, '')
                    .trim();
                
                // 验证活动名称
                const isHeaderRow = /^(序号|活动名称|加分|姓名|学号|部门|学院|系别|单位|组织)\s*[：:]/;
                const isTooShort = cleanedName.length < 2;
                const isOnlyNumbers = /^\d+$/.test(cleanedName);
                const isOnlyPunctuation = /^[^\u4e00-\u9fa5a-zA-Z0-9]+$/.test(cleanedName);
                
                if (!isHeaderRow.test(cleanedName) && !isTooShort && !isOnlyNumbers && !isOnlyPunctuation && !isNaN(score) && score > 0) {
                    const activityRecord = Object.assign({}, currentRecord); 
                    activityRecord['所参加的活动及担任角色'] = cleanedName; 
                    activityRecord['加分'] = score; 
                    records.push(activityRecord);
                    processedCount++;
                    console.log(`行 ${index + 1}: 添加活动记录成功`);
                } else {
                    console.log(`行 ${index + 1}: 验证失败，跳过`);
                    skippedCount++;
                }
                return; 
            } else {
                console.log(`行 ${index + 1}: 序号格式匹配失败`);
            }
        }
        
        // 处理普通格式的活动行
        // 改进的分数匹配正则表达式，支持更多格式
        const scorePatterns = [
            /([\d.]+)\s*$/,           // 行末分数
            /[\s\t]+([\d.]+)\s*$/,    // 制表符或空格分隔的分数
            /[：:]\s*([\d.]+)\s*$/,   // 冒号分隔的分数
            /[（(]\s*([\d.]+)\s*[）)]\s*$/, // 括号中的分数
        ];
        
        let score = null;
        let activityName = line;
        let matchedPattern = null;
        
        // 尝试多种分数匹配模式
        for (const pattern of scorePatterns) {
            const scoreMatch = line.match(pattern);
            if (scoreMatch) {
                score = parseFloat(scoreMatch[1]);
                if (!isNaN(score) && score > 0) {
                    // 移除分数部分，获取活动名称
                    activityName = line.replace(pattern, '').trim();
                    matchedPattern = pattern.toString();
                    break;
                }
            }
        }
        
        if (score && !isNaN(score) && score > 0) {
            console.log(`行 ${index + 1}: 分数匹配成功，模式: ${matchedPattern}, 分数: ${score}`);
            console.log(`行 ${index + 1}: 原始活动名称: "${activityName}"`);
            
            // 改进的序号处理逻辑
            // 支持更多序号格式：1. 1、 1) 1） (1) （1）等
            const numberPatterns = [
                /^\d{1,2}\.\s/,      // 1. 
                /^\d{1,2}、\s/,      // 1、
                /^\d{1,2}\)\s/,      // 1)
                /^\d{1,2}）\s/,      // 1）
                /^\(\d{1,2}\)\s/,    // (1)
                /^（\d{1,2}）\s/,    // （1）
                /^\d{1,2}\s/,        // 1 (单个数字后跟空格)
            ];
            
            let removedNumber = false;
            for (const pattern of numberPatterns) {
                if (pattern.test(activityName)) {
                    activityName = activityName.replace(pattern, '');
                    removedNumber = true;
                    console.log(`行 ${index + 1}: 移除序号，模式: ${pattern.toString()}`);
                    break;
                }
            }
            
            console.log(`行 ${index + 1}: 移除序号后的活动名称: "${activityName}"`);
            
            // 清理活动名称
            activityName = activityName
                .replace(/^[^\u4e00-\u9fa5a-zA-Z0-9]*/, '') // 移除开头的非文字数字字符
                .replace(/[^\u4e00-\u9fa5a-zA-Z0-9\s\-_()（）【】\[\]""''，,。.！!？?]*$/, '') // 移除结尾的无效字符
                .trim();
            
            console.log(`行 ${index + 1}: 最终活动名称: "${activityName}"`);
            
            // 放宽过滤条件，只过滤明显的表头行
            const isHeaderRow = /^(序号|活动名称|加分|姓名|学号|部门|学院|系别|单位|组织)\s*[：:]/;
            const isTooShort = activityName.length < 2;
            const isOnlyNumbers = /^\d+$/.test(activityName);
            const isOnlyPunctuation = /^[^\u4e00-\u9fa5a-zA-Z0-9]+$/.test(activityName);
            const isTotalRow = /^(合计|总计|总分|小计|总计得分|同意加分)/.test(activityName);
            
            if (!isHeaderRow.test(activityName) && !isTooShort && !isOnlyNumbers && !isOnlyPunctuation && !isTotalRow) {
                const activityRecord = Object.assign({}, currentRecord); 
                activityRecord['加分'] = score; 
                activityRecord['所参加的活动及担任角色'] = activityName; 
                records.push(activityRecord);
                processedCount++;
                console.log(`行 ${index + 1}: 添加活动记录成功: "${activityName}" (${score}分)`);
            } else {
                console.log(`行 ${index + 1}: 验证失败，跳过。原因:`, {
                    isHeaderRow: isHeaderRow.test(activityName),
                    isTooShort,
                    isOnlyNumbers,
                    isOnlyPunctuation
                });
                skippedCount++;
            }
        } else { 
            // 备用匹配模式：活动名称 + 空格/制表符 + 分数
            console.log(`行 ${index + 1}: 尝试备用匹配模式`);
            const scoreMatch = line.match(/(.+?)[\s\t]+(\d+(?:\.\d+)?)$/); 
            if (scoreMatch && scoreMatch.length >= 3) { 
                let activityName = scoreMatch[1].trim();
                const score = parseFloat(scoreMatch[2]);
                
                console.log(`行 ${index + 1}: 备用模式匹配成功，活动名称: "${activityName}", 分数: ${score}`);
                
                // 清理活动名称
                activityName = activityName
                    .replace(/^[^\u4e00-\u9fa5a-zA-Z0-9]*/, '')
                    .replace(/[^\u4e00-\u9fa5a-zA-Z0-9\s\-_()（）【】\[\]""''，,。.！!？?]*$/, '')
                    .trim();
                
                // 验证活动名称
                const isHeaderRow = /^(序号|活动名称|加分|姓名|学号|部门|学院|系别|单位|组织)\s*[：:]/;
                const isTooShort = activityName.length < 2;
                const isOnlyNumbers = /^\d+$/.test(activityName);
                const isOnlyPunctuation = /^[^\u4e00-\u9fa5a-zA-Z0-9]+$/.test(activityName);
                const isTotalRow = /^(合计|总计|总分|小计|总计得分|同意加分)/.test(activityName);
                
                if (!isHeaderRow.test(activityName) && !isTooShort && !isOnlyNumbers && !isOnlyPunctuation && !isTotalRow && !isNaN(score) && score > 0) {
                    const activityRecord = Object.assign({}, currentRecord); 
                    activityRecord['所参加的活动及担任角色'] = activityName; 
                    activityRecord['加分'] = score; 
                    records.push(activityRecord);
                    processedCount++;
                    console.log(`行 ${index + 1}: 备用模式添加活动记录成功`);
                } else {
                    console.log(`行 ${index + 1}: 备用模式验证失败，跳过`);
                    skippedCount++;
                }
            } else {
                console.log(`行 ${index + 1}: 所有匹配模式都失败，跳过`);
                skippedCount++;
            }
        } 
    });
    
    console.log(`parseActivityLines 处理完成，成功处理: ${processedCount} 个，跳过: ${skippedCount} 个`);
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
        // 改进学号提取逻辑，支持更多格式
        const sidIndex = text.search(/学号/);
        if (sidIndex !== -1) {
            const ctx = text.slice(Math.max(0, sidIndex), sidIndex + 150);
            // 支持更多学号格式
            const fieldMatches = [
                /学号[^\d\n\r]*([0-9]{6,20})/,
                /学号[：:]\s*([0-9]{6,20})/,
                /学号\s+([0-9]{6,20})/,
                /学号[^\d\n\r]*([0-9]{8,12})/,
                /学号[：:]\s*([0-9]{8,12})/,
                /学号\s+([0-9]{8,12})/
            ];
            
            for (const pattern of fieldMatches) {
                const fieldMatch = ctx.match(pattern);
                if (fieldMatch && fieldMatch[1]) {
                    studentId = fieldMatch[1];
                    console.log('parseSpecialTableFormat - 学号提取成功:', studentId, '模式:', pattern);
                    break;
                }
            }
        }
        
        // 如果还是未知，尝试全局搜索学号
        if (studentId === '未知') {
            // 搜索所有可能的学号格式
            const allStudentIdPatterns = [
                /\b([0-9]{8,12})\b/g,  // 8-12位数字
                /\b([0-9]{6,20})\b/g   // 6-20位数字
            ];
            
            for (const pattern of allStudentIdPatterns) {
                const matches = text.match(pattern);
                if (matches && matches.length > 0) {
                    // 过滤掉明显不是学号的数字（如年份、分数等）
                    const validStudentIds = matches.filter(id => {
                        const num = parseInt(id);
                        // 排除年份（1900-2030）
                        if (num >= 1900 && num <= 2030) return false;
                        // 排除分数（0-10）
                        if (num >= 0 && num <= 10) return false;
                        // 排除太短的学号
                        if (id.length < 6) return false;
                        return true;
                    });
                    
                    if (validStudentIds.length === 1) {
                        studentId = validStudentIds[0];
                        console.log('parseSpecialTableFormat - 全局学号提取成功:', studentId);
                        break;
                    } else if (validStudentIds.length > 1) {
                        // 如果有多个候选，选择最长的
                        studentId = validStudentIds.reduce((a, b) => a.length > b.length ? a : b);
                        console.log('parseSpecialTableFormat - 多个学号候选，选择最长的:', studentId);
                        break;
                    }
                }
            }
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
        
        // 改进的活动匹配逻辑，支持多种格式，避免重复提取
        console.log('parseSpecialTableFormat - 开始解析活动，文本长度:', text.length);
        console.log('parseSpecialTableFormat - 文本预览:', text.substring(0, 200)); 
        
        // 记录已匹配的位置，避免重复提取
        const matchedPositions = new Set();
        
        // 方法1: 匹配带序号的格式 (1. 活动名称 分数) - 最高优先级
        const numberedPattern = /^\d+[.\s、）\)]\s*(.+?)\s+(\d+(?:\.\d+)?)/gm;
        let match;
        while ((match = numberedPattern.exec(text)) !== null) { 
            let activityName = match[1].trim(); 
            const score = parseFloat(match[2]); 
            const matchStart = match.index;
            const matchEnd = matchStart + match[0].length;
            
            console.log('parseSpecialTableFormat - 序号格式匹配:', activityName, score, '位置:', matchStart, '-', matchEnd);
            
            if (activityName.length > 1 && !isNaN(score) && score > 0 && score <= 10 && 
                !/加分|活动名称|序号|姓名|学号|部门|学院|系别|单位|组织|合计|总计|总分/.test(activityName)) {
                console.log('parseSpecialTableFormat - 添加序号格式活动:', activityName);
                activities.push({ name: activityName, score, priority: 1 }); 
                
                // 标记这个位置已经被匹配
                for (let i = matchStart; i < matchEnd; i++) {
                    matchedPositions.add(i);
                }
            }
        }
        
        // 方法2: 匹配冒号格式 (活动名称：分数) - 第二优先级
        const colonPattern = /([^：\n\r]+)：\s*(\d+(?:\.\d+)?)/g;
        while ((match = colonPattern.exec(text)) !== null) { 
            const matchStart = match.index;
            const matchEnd = matchStart + match[0].length;
            
            // 检查这个位置是否已经被匹配
            let alreadyMatched = false;
            for (let i = matchStart; i < matchEnd; i++) {
                if (matchedPositions.has(i)) {
                    alreadyMatched = true;
                    break;
                }
            }
            
            if (alreadyMatched) {
                console.log('parseSpecialTableFormat - 跳过已匹配的冒号格式:', match[0]);
                continue;
            }
            
            let activityName = match[1].trim(); 
            const score = parseFloat(match[2]); 
            console.log('parseSpecialTableFormat - 冒号格式匹配:', activityName, score);
            
            if (activityName.length > 1 && !isNaN(score) && score > 0 && score <= 10 && 
                !/加分|活动名称|序号|姓名|学号|部门|学院|系别|单位|组织|合计|总计|总分/.test(activityName)) {
                console.log('parseSpecialTableFormat - 添加冒号格式活动:', activityName);
                activities.push({ name: activityName, score, priority: 2 }); 
                
                // 标记这个位置已经被匹配
                for (let i = matchStart; i < matchEnd; i++) {
                    matchedPositions.add(i);
                }
            }
        }
        
        // 方法3: 匹配括号格式 (活动名称（分数）) - 第三优先级
        const bracketPattern = /([^（\n\r]+)（\s*(\d+(?:\.\d+)?)\s*）/g;
        while ((match = bracketPattern.exec(text)) !== null) { 
            const matchStart = match.index;
            const matchEnd = matchStart + match[0].length;
            
            // 检查这个位置是否已经被匹配
            let alreadyMatched = false;
            for (let i = matchStart; i < matchEnd; i++) {
                if (matchedPositions.has(i)) {
                    alreadyMatched = true;
                    break;
                }
            }
            
            if (alreadyMatched) {
                console.log('parseSpecialTableFormat - 跳过已匹配的括号格式:', match[0]);
                continue;
            }
            
            let activityName = match[1].trim(); 
            const score = parseFloat(match[2]); 
            console.log('parseSpecialTableFormat - 括号格式匹配:', activityName, score);
            
            if (activityName.length > 1 && !isNaN(score) && score > 0 && score <= 10 && 
                !/加分|活动名称|序号|姓名|学号|部门|学院|系别|单位|组织|合计|总计|总分/.test(activityName)) {
                console.log('parseSpecialTableFormat - 添加括号格式活动:', activityName);
                activities.push({ name: activityName, score, priority: 3 }); 
                
                // 标记这个位置已经被匹配
                for (let i = matchStart; i < matchEnd; i++) {
                    matchedPositions.add(i);
                }
            }
        }
        
        // 方法4: 匹配空格分隔格式 (活动名称 分数) - 最低优先级，只匹配未被其他方法匹配的内容
        const spacePattern = /([^0-9\n\r]+?)\s+(\d+(?:\.\d+)?)(?=\s|$)/g;
        while ((match = spacePattern.exec(text)) !== null) { 
            const matchStart = match.index;
            const matchEnd = matchStart + match[0].length;
            
            // 检查这个位置是否已经被匹配
            let alreadyMatched = false;
            for (let i = matchStart; i < matchEnd; i++) {
                if (matchedPositions.has(i)) {
                    alreadyMatched = true;
                    break;
                }
            }
            
            if (alreadyMatched) {
                console.log('parseSpecialTableFormat - 跳过已匹配的空格格式:', match[0]);
                continue;
            }
            
            let activityName = match[1].trim(); 
            const score = parseFloat(match[2]); 
            console.log('parseSpecialTableFormat - 空格格式匹配:', activityName, score);
            
            // 过滤掉明显的非活动内容
            if (activityName.length > 1 && !isNaN(score) && score > 0 && score <= 10 && 
                !/加分|活动名称|序号|姓名|学号|部门|学院|系别|单位|组织|总计|总分|合计|小计|总计得分|同意加分/.test(activityName) &&
                !/^\d+$/.test(activityName) && // 不是纯数字
                !/^[^\u4e00-\u9fa5]+$/.test(activityName)) { // 包含中文
                console.log('parseSpecialTableFormat - 添加空格格式活动:', activityName);
                activities.push({ name: activityName, score, priority: 4 }); 
                
                // 标记这个位置已经被匹配
                for (let i = matchStart; i < matchEnd; i++) {
                    matchedPositions.add(i);
                }
            }
        }
        
        // 智能去重：优先保留更完整的活动名称（包含年份的版本）
        const uniqueActivities = [];
        const seen = new Map(); // 使用Map来存储活动名称和对应的活动对象
        
        activities.forEach(activity => {
            const key = activity.name + '|' + activity.score;
            
            if (!seen.has(key)) {
                // 第一次遇到这个活动，直接添加
                seen.set(key, activity);
                uniqueActivities.push(activity);
            } else {
                // 已经存在这个活动，比较哪个更完整
                const existing = seen.get(key);
                const existingName = existing.name;
                const newName = activity.name;
                
                // 优先保留包含年份的版本
                const existingHasYear = /\d{4}/.test(existingName);
                const newHasYear = /\d{4}/.test(newName);
                
                if (newHasYear && !existingHasYear) {
                    // 新版本包含年份，旧版本不包含，替换
                    console.log('parseSpecialTableFormat - 替换为更完整的活动名称:', existingName, '->', newName);
                    const index = uniqueActivities.indexOf(existing);
                    uniqueActivities[index] = activity;
                    seen.set(key, activity);
                } else if (newHasYear === existingHasYear) {
                    // 都包含年份或都不包含年份，优先保留优先级更高的（数字更小的）
                    if (activity.priority < existing.priority) {
                        console.log('parseSpecialTableFormat - 替换为优先级更高的活动:', existingName, '->', newName);
                        const index = uniqueActivities.indexOf(existing);
                        uniqueActivities[index] = activity;
                        seen.set(key, activity);
                    }
                }
                // 如果旧版本包含年份而新版本不包含，保持旧版本
            }
        });
        
        console.log('parseSpecialTableFormat - 最终活动数量:', uniqueActivities.length);
        
        if (uniqueActivities.length > 0) {
            uniqueActivities.forEach(activity => {
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
            const candidates = (text.match(/[\u4e00-\u9fa5·]{2,4}/g) || []).filter(t => !/姓名|学号|部门|学院|性别|民族|政治|出生|班级|专业|活动|角色|分|总分|总计/.test(t));
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
                        
                        // 使用新的合计比对功能
                        if (window.reviewRecordsWithTotal && records.length > 0) {
                            const reviewResult = reviewRecordsWithTotal(records, text);
                            // 将审查结果添加到全局变量中
                            if (reviewResult && reviewResult.reviewResults) {
                                reviewResults.push(...reviewResult.reviewResults);
                            }
                        } else {
                            // 回退到原来的审查方式
                            records.forEach(record => {
                                console.log('处理记录:', record);
                                const reviewResult = reviewRecord(record);
                                reviewResults.push(reviewResult);
                            });
                        }
                        
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
                    
                    // 对于Excel文件，也进行合计比对（如果有原始文本的话）
                    if (window.reviewRecordsWithTotal && jsonData.length > 0) {
                        // 尝试从Excel中提取文本内容用于合计比对
                        const excelText = XLSX.utils.sheet_to_txt(worksheet);
                        const reviewResult = reviewRecordsWithTotal(jsonData, excelText);
                        if (reviewResult && reviewResult.reviewResults) {
                            reviewResults.push(...reviewResult.reviewResults);
                        }
                    } else {
                        // 回退到原来的审查方式
                        jsonData.forEach(row => {
                            const result = reviewRecord(row);
                            reviewResults.push(result);
                        });
                    }
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

