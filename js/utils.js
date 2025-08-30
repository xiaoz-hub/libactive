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
