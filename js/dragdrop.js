function handleFileSelect(e) {
    const files = Array.from(e.target.files);
    processFiles(files);
    e.target.value = '';
}

function handleDragOver(e) {
    e.preventDefault();
    dropArea.classList.add('border-primary');
}

function handleDrop(e) {
    e.preventDefault();
    dropArea.classList.remove('border-primary');
    const files = Array.from(e.dataTransfer.files);
    processFiles(files);
}

function showDragOverlay(e) {
    if (e.dataTransfer && Array.from(e.dataTransfer.types || []).includes('Files')) {
        e.preventDefault();
        dragOverlay.classList.remove('hidden');
    }
}

function maybeHideDragOverlay(e) {
    if (e.target === document || e.clientX <= 0 || e.clientY <= 0 || e.clientX >= window.innerWidth || e.clientY >= window.innerHeight) {
        dragOverlay.classList.add('hidden');
    }
}

function handleGlobalDrop(e) {
    e.preventDefault();
    dragOverlay.classList.add('hidden');
    const files = Array.from(e.dataTransfer.files || []);
    if (files.length > 0) processFiles(files);
}

function processFiles(files) {
    const validFiles = files.filter(file => {
        const ext = file.name.split('.').pop().toLowerCase();
        return ['xlsx', 'xls', 'csv', 'doc', 'docx'].includes(ext);
    });
    const existingKeys = new Set(uploadedFiles.map(f => `${f.name}|${f.size}|${f.lastModified || ''}`));
    const deduped = [];
    let skipped = 0;
    for (const f of validFiles) {
        const key = `${f.name}|${f.size}|${f.lastModified || ''}`;
        if (!existingKeys.has(key)) { existingKeys.add(key); deduped.push(f); }
        else skipped += 1;
    }
    if (deduped.length > 0) {
        uploadedFiles = [...uploadedFiles, ...deduped];
        updateFileList();
        // 只要有文件就启用开始审查按钮，没有活动规则时点击会显示弹窗
        startReviewBtn.disabled = false;
        const msg = skipped > 0 ? `成功上传 ${deduped.length} 个文件（已跳过重复 ${skipped} 个）` : `成功上传 ${deduped.length} 个文件`;
        showNotification(msg);
    } else {
        if (skipped > 0) showNotification(`全部为重复文件，已跳过 ${skipped} 个`, 'info');
        else showNotification('请上传Excel、CSV或Word格式的文件', 'error');
    }
}

