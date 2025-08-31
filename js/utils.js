function showNotification(message, type = 'success') {
    if (notificationTimeout) clearTimeout(notificationTimeout);
    let notification = document.getElementById('notification');
    if (!notification) {
        notification = document.createElement('div');
        notification.id = 'notification';
        notification.className = 'fixed top-4 right-4 z-50 px-4 py-3 rounded-md shadow-lg transform transition-all duration-300';
        document.body.appendChild(notification);
    }
    notification.textContent = message;
    notification.className = 'fixed top-4 right-4 z-50 px-4 py-3 rounded-md shadow-lg transform transition-all duration-300';
    if (type === 'success') notification.classList.add('bg-success', 'text-white');
    else if (type === 'error') notification.classList.add('bg-danger', 'text-white');
    else if (type === 'warning') notification.classList.add('bg-warning', 'text-white', 'font-semibold', 'border-2', 'border-yellow-400');
    else if (type === 'info') notification.classList.add('bg-info', 'text-white');
    notification.style.opacity = '0';
    notification.style.transform = 'translateX(100%)';
    notification.style.display = 'block';
    setTimeout(() => { notification.style.opacity = '1'; notification.style.transform = 'translateX(0)'; }, 10);
    notificationTimeout = setTimeout(() => {
        notification.style.opacity = '0';
        notification.style.transform = 'translateX(100%)';
        setTimeout(() => { notification.style.display = 'none'; }, 300);
    }, 3000);
}

function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

function levenshteinDistance(a, b) {
    if (a.length === 0) return b.length;
    if (b.length === 0) return a.length;
    const matrix = [];
    for (let i = 0; i <= b.length; i++) matrix[i] = [i];
    for (let j = 0; j <= a.length; j++) matrix[0][j] = j;
    for (let i = 1; i <= b.length; i++) {
        for (let j = 1; j <= a.length; j++) {
            if (b.charAt(i-1) === a.charAt(j-1)) matrix[i][j] = matrix[i-1][j-1];
            else matrix[i][j] = Math.min(matrix[i-1][j-1] + 1, matrix[i][j-1] + 1, matrix[i-1][j] + 1);
        }
    }
    return matrix[b.length][a.length];
}


// 统一破折号/连接号等变体，避免因字符不同导致匹配失败
function normalizeActivityName(value) {
    if (value == null) return '';
    return String(value)
        // 各类破折号/连接号/全角减号 → 半角短横线
        .replace(/[\u2014\u2013\u2010\u2012\uFF0D\u2212\uFE63\u02D7\u2043\u2053\-]+/g, '-')
        // 波浪线等分隔符也统一
        .replace(/[~〜～﹏﹋﹌]/g, '-')
        // 去除横线两侧多余空白并压缩空白
        .replace(/\s+/g, ' ')
        .replace(/\s*-\s*/g, '-')
        .trim();
}

// 智能分割中文活动名称中的词汇
function extractChineseWords(text) {
    if (!text) return [];
    
    const words = [];
    
    // 常见活动相关词汇
    const commonWords = [
        '活动', '工作人员', '活动人员', '副部', '部长', '副部长', '干事', '成员', '负责人',
        '草坪', '图书馆', '期刊', '人事', '知识', '竞赛', '运动会', '猜灯谜',
        '寻宝', '大赛', '盛宴', '青春', '智慧', '闪耀', '飞扬', '迎新', '招新',
        '嘉年华', '教育', '入馆', '新生'
    ];
    
    // 查找包含的常见词汇
    commonWords.forEach(word => {
        if (text.includes(word)) {
            words.push(word);
        }
    });
    
    // 按2-4个字符分割，提取可能的词汇
    for (let i = 0; i < text.length - 1; i++) {
        for (let len = 2; len <= 4 && i + len <= text.length; len++) {
            const word = text.substring(i, i + len);
            if (word.length >= 2) {
                words.push(word);
            }
        }
    }
    
    // 添加一些特殊的分割逻辑
    // 处理"活动人员"、"工作人员"等组合词
    if (text.includes('活动人员')) {
        words.push('活动人员');
    }
    if (text.includes('工作人员')) {
        words.push('工作人员');
    }
    if (text.includes('草坪活动')) {
        words.push('草坪活动');
    }
    if (text.includes('猜灯谜活动')) {
        words.push('猜灯谜活动');
    }
    
    return words;
}
