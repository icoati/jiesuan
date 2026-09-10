/**
 * ====================================================
 * 心血管内科报表生成脚本（自包含版）
 *
 * 使用方法：
 *   1. 把两个系统导出的 xlsx 文件放到本脚本同一文件夹
 *   2. 双击「一键统计.command」或在终端运行：node 统计报表.js
 *   3. 生成的报表会出现在同一文件夹
 * ====================================================
 */

const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// === 配置：从脚本所在文件夹读取 ===
const SCRIPT_DIR = __dirname;
const OUTPUT_NAME = '统计总表.xlsx';

// === 年龄转年龄段 ===
function convertAgeGroup(age) {
    const n = parseInt(age);
    if (isNaN(n) || n <= 0) return '';
    if (n <= 40) return '40岁以下';
    if (n <= 50) return '41-50岁';
    if (n <= 60) return '51-60岁';
    if (n <= 70) return '61-70岁';
    return '71岁以上';
}

// === 自动从脚本所在文件夹查找目标 xlsx 文件 ===
function findSourceFiles() {
    const files = fs.readdirSync(SCRIPT_DIR).filter(f => {
        return f.endsWith('.xlsx') && f !== OUTPUT_NAME && !f.startsWith('~$');
    });

    let srcFile = null;
    let staffFile = null;

    for (const file of files) {
        const filePath = path.join(SCRIPT_DIR, file);
        try {
            const wb = XLSX.readFile(filePath);
            const names = wb.SheetNames;

            if (!srcFile && names.includes('答卷记录')) {
                srcFile = { path: filePath, name: file };
            }
            if (!staffFile && names.includes('项目人员')) {
                staffFile = { path: filePath, name: file };
            }
        } catch (e) {
            // 跳过无法读取的文件
        }
        if (srcFile && staffFile) break;
    }

    return { srcFile, staffFile };
}

// === 读取答卷记录 ===
function readSource(filePath) {
    const wb = XLSX.readFile(filePath);
    const ws = wb.Sheets['答卷记录'];
    if (!ws) {
        console.error('❌ 文件中未找到「答卷记录」工作表');
        process.exit(1);
    }
    return XLSX.utils.sheet_to_json(ws, { header: 1 });
}

// === 读取项目人员列表 ===
function readStaff(filePath) {
    const wb = XLSX.readFile(filePath);
    const ws = wb.Sheets['项目人员'];
    if (!ws) {
        console.error('❌ 文件中未找到「项目人员」工作表');
        process.exit(1);
    }
    return XLSX.utils.sheet_to_json(ws, { header: 1 });
}

// === 构建 Sheet1: 统计汇总 ===
function buildSummary(staffData) {
    const provMap = {};
    let total = 0;

    for (let i = 1; i < staffData.length; i++) {
        const row = staffData[i];
        const taskName = String(row[0] || '');
        const province = String(row[5] || '');
        const count = parseInt(row[7]) || 0;

        if (taskName.includes('心血管内科')) {
            if (!provMap[province]) provMap[province] = 0;
            provMap[province] += count;
            total += count;
        }
    }

    const provinces = Object.keys(provMap).sort();
    const result = [['医院所在省', '药品规格', '任务数量']];
    for (const prov of provinces) {
        const spec = prov === '上海市' ? '片剂' : '滴丸';
        result.push([prov, spec, provMap[prov]]);
    }
    result.push(['总计', null, total]);
    return result;
}

// === 构建 Sheet2: 人员维度 ===
function buildStaffSheet(staffData) {
    const staff = [];
    for (let i = 1; i < staffData.length; i++) {
        const row = staffData[i];
        const taskName = String(row[0] || '');
        if (taskName.includes('心血管内科')) {
            staff.push({
                任务名称: taskName,
                姓名: String(row[1] || ''),
                手机号: '',
                医院: String(row[2] || ''),
                科室: String(row[3] || ''),
                职称: String(row[4] || ''),
                医院所在省: String(row[5] || ''),
                医院所在市: String(row[6] || ''),
                开工状态: '已开工',
                资质是否通过: '是',
                任务完成情况: parseInt(row[7]) || 0,
                更新时间: String(row[8] || '')
            });
        }
    }

    staff.sort((a, b) => b.更新时间.localeCompare(a.更新时间));

    const header = ['任务名称', '姓名', '手机号', '医院', '科室', '职称',
        '医院所在省', '医院所在市', '开工状态', '资质是否通过', '任务完成情况', '更新时间'];

    const result = [header];
    for (const s of staff) {
        result.push([
            s.任务名称, s.姓名, s.手机号, s.医院, s.科室, s.职称,
            s.医院所在省, s.医院所在市, s.开工状态, s.资质是否通过,
            s.任务完成情况, s.更新时间
        ]);
    }
    return result;
}

// === 构建 Sheet3: 案例维度 ===
function buildCaseSheet(sourceData, staffData) {
    const staffSet = new Set();
    for (let i = 1; i < staffData.length; i++) {
        const row = staffData[i];
        const taskName = String(row[0] || '');
        if (taskName.includes('心血管内科')) {
            staffSet.add(String(row[1] || '').trim());
        }
    }

    const header = [
        '序号', '任务名称', '项目人员', '手机号', '医院', '医院所在省份',
        '医院所在市', '结算单号', '答卷标题', '审核状态', '审核未通过原因',
        '提交答卷时间', '1、姓名（首字母缩写）：', '2、性别：', '3、年龄：（岁）',
        '4、诊断日期:', '5、初诊 / 复诊', '6、临床表现（可多选）：',
        '7、疾病诊断（可多选）：', '8、既往史（可多选）：', '33、其他改善：'
    ];

    const result = [header];
    let seqNum = 1;

    for (let i = 1; i < sourceData.length; i++) {
        const row = sourceData[i];
        const staffName = String(row[2] || '').trim();

        // 导出所有记录（含审核通过和不通过），在「审核状态」列区分
        if (staffSet.has(staffName)) {
            result.push([
                seqNum,
                String(row[1] || ''),
                staffName,
                '',
                String(row[3] || ''),
                String(row[4] || ''),
                String(row[5] || ''),
                '',
                String(row[8] || ''),
                String(row[12] || ''),
                String(row[13] || ''),
                String(row[14] || ''),
                String(row[15] || ''),
                String(row[16] || ''),
                convertAgeGroup(row[17]),
                String(row[18] || ''),
                String(row[19] || ''),
                String(row[20] || ''),
                String(row[21] || ''),
                String(row[22] || ''),
                String(row[33] || '')
            ]);
            seqNum++;
        }
    }
    return result;
}

// === 设置工作表样式 ===
function applyStyle(wb, sheetName, data) {
    const ws = wb.Sheets[sheetName];
    const headerRange = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    const colCount = headerRange.e.c - headerRange.s.c + 1;

    for (let c = 0; c < colCount; c++) {
        const cellRef = XLSX.utils.encode_cell({ r: 0, c: c });
        if (ws[cellRef]) {
            ws[cellRef].s = {
                font: { bold: true, sz: 11 },
                fill: { fgColor: { rgb: 'C8DCF0' } },
                alignment: { horizontal: 'center', vertical: 'center' }
            };
        }
    }

    ws['!cols'] = [];
    for (let c = 0; c < colCount; c++) {
        let maxLen = 0;
        for (let r = 0; r < data.length && r < 50; r++) {
            const val = String(data[r][c] || '');
            let len = 0;
            for (const ch of val) {
                len += ch.charCodeAt(0) > 127 ? 2 : 1;
            }
            if (len > maxLen) maxLen = len;
        }
        ws['!cols'].push({ wch: Math.min(Math.max(maxLen + 2, 8), 40) });
    }
}

// === 主流程 ===
function main() {
    console.log('='.repeat(50));
    console.log('  心血管内科报表生成工具');
    console.log('='.repeat(50));
    console.log('📁 工作目录：' + SCRIPT_DIR);

    console.log('\n🔍 正在查找文件夹中的源文件...');
    const { srcFile, staffFile } = findSourceFiles();

    if (!srcFile) {
        console.error('❌ 未找到包含「答卷记录」工作表的 xlsx 文件');
        console.error('   请将系统导出的答卷记录表放到本文件夹后重试。');
        process.exit(1);
    }
    if (!staffFile) {
        console.error('❌ 未找到包含「项目人员」工作表的 xlsx 文件');
        console.error('   请将系统导出的项目人员列表放到本文件夹后重试。');
        process.exit(1);
    }

    console.log('📖 正在读取数据...');
    const sourceData = readSource(srcFile.path);
    console.log(`   ✓ 答卷记录：${srcFile.name}（${sourceData.length - 1} 条记录）`);

    const staffData = readStaff(staffFile.path);
    console.log(`   ✓ 项目人员：${staffFile.name}（${staffData.length - 1} 位人员）`);

    console.log('\n🔧 正在生成报表...');

    const summaryData = buildSummary(staffData);
    const staffSheetData = buildStaffSheet(staffData);
    const caseSheetData = buildCaseSheet(sourceData, staffData);

    console.log(`   ✓ 统计汇总：${summaryData.length - 2} 个省份`);
    console.log(`   ✓ 人员维度：${staffSheetData.length - 1} 位人员`);
    console.log(`   ✓ 案例维度：${caseSheetData.length - 1} 条案例`);

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(summaryData), '统计汇总');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(staffSheetData), '人员维度');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(caseSheetData), '案例维度');

    applyStyle(wb, '统计汇总', summaryData);
    applyStyle(wb, '人员维度', staffSheetData);
    applyStyle(wb, '案例维度', caseSheetData);

    // 统计汇总总计行合并 B、C 列
    const summaryWs = wb.Sheets['统计汇总'];
    const totalRow = summaryData.length - 1;
    if (!summaryWs['!merges']) summaryWs['!merges'] = [];
    summaryWs['!merges'].push({ s: { r: totalRow, c: 1 }, e: { r: totalRow, c: 2 } });

    const outputPath = path.join(SCRIPT_DIR, OUTPUT_NAME);
    XLSX.writeFile(wb, outputPath);

    console.log(`\n✅ 报表已生成：${outputPath}`);
    console.log('   包含工作表：统计汇总 / 人员维度 / 案例维度');
}

main();
