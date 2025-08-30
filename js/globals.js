// 全局变量
let uploadedFiles = [];
let activityRules = [];
let departmentRules = [];
let studentIdRules = [];
let reviewResults = [];
let currentFilter = 'all';
let currentExportFilter = 'all';

// DOM 元素
const dropArea = document.getElementById('dropArea');
const fileInput = document.getElementById('fileInput');
const browseBtn = document.getElementById('browseBtn');
const uploadedFilesDiv = document.getElementById('uploadedFiles');
const fileList = document.getElementById('fileList');
const startReviewBtn = document.getElementById('startReviewBtn');
const exportResultsBtn = document.getElementById('exportResultsBtn');
const clearResultsBtn = document.getElementById('clearResultsBtn');
const reviewStatus = document.getElementById('reviewStatus');
const progressBar = document.getElementById('progressBar');
const progressPercentage = document.getElementById('progressPercentage');
const resultsTable = document.getElementById('resultsTable');
const resultFilter = document.getElementById('resultFilter');
const resultStats = document.getElementById('resultStats');
const totalRecords = document.getElementById('totalRecords');
const passedRecords = document.getElementById('passedRecords');
const failedRecords = document.getElementById('failedRecords');
const helpBtn = document.getElementById('helpBtn');
const helpModal = document.getElementById('helpModal');
const closeHelpModal = document.getElementById('closeHelpModal');
const activityRuleModal = document.getElementById('activityRuleModal');
const closeActivityRuleModal = document.getElementById('closeActivityRuleModal');
const goToActivityRules = document.getElementById('goToActivityRules');
const studentIdRuleModal = document.getElementById('studentIdRuleModal');
const closeStudentIdRuleModal = document.getElementById('closeStudentIdRuleModal');
const goToStudentIdRules = document.getElementById('goToStudentIdRules');
const departmentRuleModal = document.getElementById('departmentRuleModal');
const closeDepartmentRuleModal = document.getElementById('closeDepartmentRuleModal');
const goToDepartmentRules = document.getElementById('goToDepartmentRules');
const themeToggle = document.getElementById('themeToggle');
const dragOverlay = document.getElementById('dragOverlay');

// 规则设置相关元素
const ruleTabs = document.querySelectorAll('.rule-tab');
const ruleContents = document.querySelectorAll('.rule-content');
const activitySelect = document.getElementById('activitySelect');
const activityName = document.getElementById('activityName');
const activityScore = document.getElementById('activityScore');
const addActivityRule = document.getElementById('addActivityRule');
const activityRulesTable = document.getElementById('activityRulesTable');
const studentIdPatternInput = document.getElementById('studentIdPattern');
const saveStudentIdRule = document.getElementById('saveStudentIdRule');
const currentStudentIdRule = document.getElementById('currentStudentIdRule');
const departmentSelect = document.getElementById('departmentSelect');
const departmentName = document.getElementById('departmentName');
const addDepartmentRule = document.getElementById('addDepartmentRule');
const departmentRulesTable = document.getElementById('departmentRulesTable');
const filterBtns = document.querySelectorAll('.filter-btn');

// 查询相关元素
const queryDepartment = document.getElementById('queryDepartment');
const queryStudentPrefix = document.getElementById('queryStudentPrefix');
const applyQueryBtn = document.getElementById('applyQueryBtn');
const resetQueryBtn = document.getElementById('resetQueryBtn');

// 当前查询条件
let currentQuery = { department: '', studentPrefix: '' };

// 其他
let notificationTimeout = null;

