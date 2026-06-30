const fs = require('fs');
const path = require('path');
const os = require('os');
const { execSync } = require('child_process');

const isRestore = process.argv.includes('--restore');
const pluginsDir = path.join(__dirname, '.obsidian', 'plugins');

if (!fs.existsSync(pluginsDir)) {
    console.error('❌ 錯誤：找不到 .obsidian/plugins 目錄。');
    console.error('請確保將此腳本放在 Obsidian Vault 的根目錄下執行。');
    process.exit(1);
}

// 遞迴遍歷目錄
function walk(dir, callback) {
    if (!fs.existsSync(dir)) return;
    const files = fs.readdirSync(dir);
    for (const file of files) {
        const filepath = path.join(dir, file);
        const stat = fs.statSync(filepath);
        if (stat.isDirectory()) {
            walk(filepath, callback);
        } else if (stat.isFile()) {
            callback(filepath);
        }
    }
}

if (isRestore) {
    console.log('🔄 開始還原備份檔案...');
    let restoreCount = 0;
    walk(pluginsDir, (filePath) => {
        if (filePath.endsWith('.bak')) {
            const originalPath = filePath.slice(0, -4);
            try {
                fs.copyFileSync(filePath, originalPath);
                fs.unlinkSync(filePath);
                console.log(`[已還原] ${path.relative(pluginsDir, originalPath)}`);
                restoreCount++;
            } catch (err) {
                console.error(`還原失敗: ${filePath}`, err);
            }
        }
    });
    console.log(`✨ 還原完成，共還原了 ${restoreCount} 個檔案。`);
    process.exit(0);
}

// 1. 準備臨時目錄安裝 opencc-js，避免污染 Vault 的 node_modules
const tempDir = path.join(os.tmpdir(), 'obsidian-tc-converter');
if (!fs.existsSync(tempDir)) {
    fs.mkdirSync(tempDir, { recursive: true });
}

const nodeModulesPath = path.join(tempDir, 'node_modules');
if (!fs.existsSync(nodeModulesPath)) {
    console.log('⏳ 正在背景初始化簡繁轉換引擎 (opencc-js)，請稍候...');
    try {
        fs.writeFileSync(path.join(tempDir, 'package.json'), JSON.stringify({ name: 'obsidian-tc-converter', version: '1.0.0' }));
        execSync('npm install opencc-js@1.0.5 --no-audit --no-fund', { cwd: tempDir, stdio: 'ignore' });
        console.log('✅ 轉換引擎初始化成功。');
    } catch (err) {
        console.error('❌ 初始化轉換引擎失敗，請檢查網路連接或是否已安裝 Node.js：', err.message);
        process.exit(1);
    }
}

// 2. 載入轉換引擎
let convert;
try {
    const OpenCC = require(path.join(tempDir, 'node_modules', 'opencc-js'));
    convert = OpenCC.Converter({ from: 'cn', to: 'tw' });
} catch (err) {
    console.error('❌ 載入轉換引擎失敗。');
    process.exit(1);
}

// Unicode 解碼與編碼輔助函式
function decodeUnicode(str) {
    return str.replace(/\\u([0-9a-fA-F]{4})/g, (match, grp) => {
        return String.fromCharCode(parseInt(grp, 16));
    });
}

function encodeUnicode(str) {
    return str.replace(/[^\x00-\x7F]/g, (char) => {
        return '\\u' + ('0000' + char.charCodeAt(0).toString(16)).slice(-4).toUpperCase();
    });
}

console.log('🚀 開始掃描外掛目錄並進行簡繁轉換...');
let convertCount = 0;

walk(pluginsDir, (filePath) => {
    const ext = path.extname(filePath);
    const fileName = path.basename(filePath);

    // 只處理 json 語系檔與 main.js 主程式
    const isLangJson = ext === '.json' && (fileName.includes('zh') || fileName.includes('cn') || fileName.includes('lang'));
    const isMainJs = fileName === 'main.js';

    if (isLangJson || isMainJs) {
        try {
            const content = fs.readFileSync(filePath, 'utf8');
            
            // 先還原可能被轉義的 \uXXXX 簡體字
            const decodedContent = decodeUnicode(content);
            const convertedDecoded = convert(decodedContent);
            
            // 如果原本是 main.js，將轉換後的繁體中文重新編碼回大寫的 \uXXXX 格式寫回
            let finalContent;
            if (isMainJs) {
                finalContent = encodeUnicode(convertedDecoded);
            } else {
                finalContent = convertedDecoded;
            }

            if (content !== finalContent) {
                // 備份原始檔案（只在沒有備份檔時備份，防止重複執行覆蓋了正確的備份）
                const backupPath = filePath + '.bak';
                if (!fs.existsSync(backupPath)) {
                    fs.writeFileSync(backupPath, content, 'utf8');
                }

                fs.writeFileSync(filePath, finalContent, 'utf8');
                console.log(`[已轉換] ${path.relative(pluginsDir, filePath)}`);
                convertCount++;
            }

            // 如果是簡中語系檔，額外多生一份繁中語系檔 zh-tw.json
            if (isLangJson && (fileName === 'zh.json' || fileName === 'zh-cn.json' || fileName === 'zh_cn.json')) {
                const twFileName = fileName.replace(/zh-cn|zh_cn|zh/, 'zh-tw');
                const twFilePath = path.join(path.dirname(filePath), twFileName);
                if (twFilePath !== filePath && !fs.existsSync(twFilePath)) {
                    fs.writeFileSync(twFilePath, convertedDecoded, 'utf8');
                    console.log(`[已建立繁中語系檔] ${path.relative(pluginsDir, twFilePath)}`);
                }
            }
        } catch (err) {
            console.error(`❌ 處理檔案失敗: ${path.relative(pluginsDir, filePath)}`, err.message);
        }
    }
});

console.log(`\n✨ 轉換完成！共修改/建立了 ${convertCount} 個檔案。`);
console.log('💡 請重新啟動 Obsidian 或在設定中重新開關外掛，即可看到繁體中文效果。');
console.log('👉 若有開關外掛異常，可執行 \`node convert-obsidian-plugins.js --restore\` 還原。');
