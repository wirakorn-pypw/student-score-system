// ==============================
// ข้อมูลหลัก
// ==============================

// กำหนดข้อมูลเริ่มต้นหาก localStorage หรือ Google Sheets ยังไม่มีข้อมูล
let students = JSON.parse(localStorage.getItem('students')) || [
    { id: "5345", code: "5345", name: "กรวิชญ์", className: "ม.6/1" },
    { id: "5344", code: "5344", name: "กนกพร", className: "ม.6/1" },
    { id: "5346", code: "5346", name: "กุสุมาภรณ์", className: "ม.6/1" },
    { id: "5349", code: "5349", name: "ญาณิศา", className: "ม.6/1" }
];
let subjects = JSON.parse(localStorage.getItem("subjects")) || [];
let scores = JSON.parse(localStorage.getItem("scores")) || [];


// ==============================
// เปลี่ยนหน้า
// ==============================

function showPage(pageName) {
    document.querySelectorAll(".page").forEach(page => {
        page.classList.add("hidden");
    });

    document.getElementById(pageName).classList.remove("hidden");
    updateAll();
}


// ==============================
// จัดการนักเรียน
// ==============================

function addStudent() {
    const code = document.getElementById("studentCode") ? document.getElementById("studentCode").value.trim() : "";
    const name = document.getElementById("studentName").value.trim();
    const studentClass = document.getElementById("studentClass").value.trim();

    if (!name || !studentClass) {
        alert("กรุณากรอกชื่อและห้องเรียน");
        return;
    }

    students.push({
        id: Date.now(),
        code: code || "-",
        name: name,
        className: studentClass
    });

    saveData();
    if (document.getElementById("studentCode")) document.getElementById("studentCode").value = "";
    document.getElementById("studentName").value = "";
    document.getElementById("studentClass").value = "";

    updateAll();
}
// ฟังก์ชันเพิ่มนักเรียนใหม่แบบสมบูรณ์
async function addStudent(code, name, className) {
    if (!code || !name) {
        alert("กรุณากรอกรหัสและชื่อนักเรียนให้ครบถ้วน");
        return;
    }

    const cleanCode = String(code).trim();
    const cleanName = String(name).trim();
    const cleanClass = className ? String(className).trim() : "ม.6/1";

    // เช็กนักเรียนซ้ำ
    const isDuplicate = students.some(s => String(s.id).trim() === cleanCode || String(s.code).trim() === cleanCode);
    if (isDuplicate) {
        alert("รหัสนักเรียนนี้มีอยู่ในระบบแล้วครับ");
        return;
    }

    const newStudent = {
        id: cleanCode,
        code: cleanCode,
        name: cleanName,
        className: cleanClass
    };

    // 1. เพิ่มเข้าตัวแปรในระบบ
    students.push(newStudent);

    // 2. บันทึกลงความจำเบราว์เซอร์
    localStorage.setItem('students', JSON.stringify(students));

    // 3. วาดตารางแสดงผลใหม่ทันที
    if (typeof renderScoreMatrix === 'function') renderScoreMatrix();
    if (typeof calculateAndRenderSummaryScores === 'function') calculateAndRenderSummaryScores();

    alert(`เพิ่มนักเรียน ${cleanName} เรียบร้อยแล้ว!`);

    // 4. ส่งไปบันทึกลง Google Sheets (ถ้าเปิดใช้)
    if (typeof GAS_API_URL !== 'undefined' && GAS_API_URL !== "") {
        try {
            await fetch(GAS_API_URL, {
                method: "POST",
                body: JSON.stringify({
                    action: "addStudent",
                    student: newStudent
                })
            });
        } catch (e) {
            console.error("ส่งข้อมูลลง Google Sheets ไม่สำเร็จ:", e);
        }
    }
}
// ==============================
// จัดการนำเข้าไฟล์นักเรียน (CSV / XLSX / XLS)
// ==============================

function importStudentsFromFile() {
    const fileInput = document.getElementById("csvFileInput");
    const importClass = document.getElementById("importClass") ? document.getElementById("importClass").value.trim() : "";
    const file = fileInput.files[0];

    if (!importClass || !file) {
        alert("กรุณาระบุห้องเรียนและเลือกไฟล์ตารางข้อมูลครับ");
        return;
    }

    const reader = new FileReader();
    const fileName = file.name.toLowerCase();

    // กรณีเป็นไฟล์ Excel (.xlsx, .xls)
    if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
        reader.onload = function (e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            // ดึง Sheet แรกสุด
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            // แปลงตารางเป็น Array
            const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            processImportedRows(rows, importClass, fileInput);
        };
        reader.readAsArrayBuffer(file);
    } 
    // กรณีเป็นไฟล์ CSV
    else {
        reader.onload = function (e) {
            const text = e.target.result;
            const lines = text.split(/\r\n|\n/);
            const rows = lines.map(line => line.split(","));
            processImportedRows(rows, importClass, fileInput);
        };
        reader.readAsText(file, "UTF-8");
    }
}

// ฟังก์ชันประมวลผลข้อมูลแถว
function processImportedRows(rows, importClass, fileInput) {
    let importedCount = 0;

    rows.forEach((parts, index) => {
        if (!parts || parts.length === 0) return;

        const col0 = parts[0] ? String(parts[0]).trim() : "";
        const col1 = parts[1] ? String(parts[1]).trim() : "";
        const col2 = parts[2] ? String(parts[2]).trim() : "";

        // ข้ามแถว Header (ถ้ามีคำว่า เลขที่ / รหัส / ชื่อ)
        if (index === 0 && (col0.includes("เลขที่") || col1.includes("รหัส") || col2.includes("ชื่อ"))) return;

        const code = col1;
        const name = col2 || col1; // ถ้าไม่มีช่อง 3 ให้ใช้ช่อง 2 เป็นชื่อ

        if (name && name !== code) {
            students.push({
                id: Date.now() + Math.floor(Math.random() * 1000) + index,
                code: code || "-",
                name: name,
                className: importClass
            });
            importedCount++;
        }
    });

    if (importedCount > 0) {
        saveData();
        updateAll();
        fileInput.value = "";
        if (document.getElementById("importClass")) document.getElementById("importClass").value = "";
        alert(`นำเข้าข้อมูลนักเรียนห้อง ${importClass} เรียบร้อยแล้วจำนวน ${importedCount} คน`);
    } else {
        alert("ไม่พบข้อมูลที่ถูกต้องในไฟล์");
    }
}

function deleteStudent(id) {
    if (!confirm("ต้องการลบนักเรียนคนนี้หรือไม่?")) return;
    students = students.filter(student => student.id !== id);
    scores = scores.filter(score => score.studentId !== id);
    saveData();
    updateAll();
}

function renderStudents() {
    const table = document.getElementById("studentTable");
    const filterSelect = document.getElementById("classFilterSelect");
    if (!table) return;

    // ดึงรายชื่อห้องทั้งหมดที่มีอยู่แบบไม่ซ้ำกัน
    const uniqueClasses = [...new Set(students.map(s => s.className))].sort();

    // อัปเดตรายการใน Dropdown ตัวกรองโฟลเดอร์
    if (filterSelect) {
        const currentSelected = filterSelect.value || "ALL";
        filterSelect.innerHTML = `<option value="ALL">📂 แสดงทุกห้อง (${students.length} คน)</option>`;
        
        uniqueClasses.forEach(cls => {
            const count = students.filter(s => s.className === cls).length;
            const option = document.createElement("option");
            option.value = cls;
            option.textContent = `📁 ห้อง ${cls} (${count} คน)`;
            filterSelect.appendChild(option);
        });

        // คืนค่าห้องที่เคยเลือกไว้
        if (uniqueClasses.includes(currentSelected) || currentSelected === "ALL") {
            filterSelect.value = currentSelected;
        }
    }

    const selectedClass = filterSelect ? filterSelect.value : "ALL";
    
    // กรองนักเรียนตามห้องที่เลือก
    const filteredStudents = selectedClass === "ALL" 
        ? students 
        : students.filter(s => s.className === selectedClass);

    table.innerHTML = "";

    if (filteredStudents.length === 0) {
        table.innerHTML = `<tr><td colspan="5" style="text-align: center; color: #94a3b8; padding: 20px;">ไม่พบข้อมูลนักเรียนในหมวดหมู่นี้</td></tr>`;
        return;
    }

    filteredStudents.forEach((student, index) => {
        const row = document.createElement("tr");
        row.innerHTML = `
            <td>${index + 1}</td>
            <td>${student.code || "-"}</td>
            <td>${student.name}</td>
            <td><span style="background: #e2e8f0; padding: 2px 8px; border-radius: 4px; font-size: 13px;">${student.className}</span></td>
            <td><button onclick="deleteStudent(${student.id})" style="background: #ef4444; color: white; padding: 4px 8px; font-size: 12px;">ลบ</button></td>
        `;
        table.appendChild(row);
    });
}

// ฟังก์ชันลบเฉพาะห้องที่เลือกอยู่ปัจจุบันใน Folder Filter
function deleteStudentsBySelectedClass() {
    const filterSelect = document.getElementById("classFilterSelect");
    const targetClass = filterSelect ? filterSelect.value : "ALL";

    if (targetClass === "ALL") {
        alert("กรุณาเลือกห้องที่ต้องการลบจากเมนู 'เลือกดูตามห้อง/โฟลเดอร์' ก่อนครับ");
        return;
    }

    const studentsInClass = students.filter(s => s.className === targetClass);

    if (confirm(`คุณต้องการลบนักเรียนห้อง ${targetClass} ทั้งหมดจำนวน ${studentsInClass.length} คนใช่หรือไม่?\n(คะแนนสอบของนักเรียนกลุ่มนี้จะถูกลบออกด้วย)`)) {
        const studentIdsToRemove = new Set(studentsInClass.map(s => s.id));

        students = students.filter(s => s.className !== targetClass);
        scores = scores.filter(s => !studentIdsToRemove.has(s.studentId));

        saveData();
        updateAll();
        alert(`ลบข้อมูลนักเรียนห้อง ${targetClass} เรียบร้อยแล้ว`);
    }
}


// ==============================
// จัดการวิชา บทเรียน และหน่วยย่อย
// ==============================

// ==============================
// จัดการวิชา (เพิ่มระดับชั้น)
// ==============================

function addSubject() {
    const grade = document.getElementById("subjectGrade") ? document.getElementById("subjectGrade").value : "ม.4";
    const code = document.getElementById("subjectCode").value.trim();
    const name = document.getElementById("subjectName").value.trim();
    const midtermMax = Number(document.getElementById("midtermScore")?.value) || 20;
    const finalMax = Number(document.getElementById("finalScore")?.value) || 30;

    if (!code || !name) {
        alert("กรุณากรอกรหัสวิชาและชื่อวิชา");
        return;
    }

    if (midtermMax + finalMax > 100) {
        alert("คะแนนกลางภาค + ปลายภาค รวมกันเกิน 100 คะแนนไม่ได้ครับ");
        return;
    }

    subjects.push({
        id: Date.now(),
        grade: grade,
        code: code,
        name: name,
        midtermMax: midtermMax,
        finalMax: finalMax,
        units: []
    });

    saveData();
    document.getElementById("subjectCode").value = "";
    document.getElementById("subjectName").value = "";
    updateAll();
}

function renderSubjects() {
    const container = document.getElementById("subjectListContainer");
    if (!container) return;

    container.innerHTML = "";
    if (subjects.length === 0) {
        container.innerHTML = "<p style='color: #888;'>ยังไม่มีข้อมูลวิชา</p>";
        return;
    }

    subjects.forEach(subject => {
        const totalUnitQuota = 100 - ((subject.midtermMax || 0) + (subject.finalMax || 0));

        const card = document.createElement("div");
        card.style.cssText = "background: #f9fafb; border: 1px solid #e5e7eb; border-radius: 8px; padding: 15px; margin-bottom: 15px;";
        
        let unitsHtml = "";
        if (subject.units && subject.units.length > 0) {
            unitsHtml = `<div style="margin-top: 10px;">` +
                subject.units.map(u => {
                    let subHtml = "";

                    if (u.subUnits && u.subUnits.length > 0) {
                        subHtml = `<ul style="margin: 5px 0 10px 20px; font-size: 13px; color: #4b5563;">` +
                            u.subUnits.map(sub => `
                                <li>
                                    ${sub.name} <span style="color: #059669; font-weight: bold;">(คำนวณเต็ม: ${sub.maxScore} คะแนน)</span>
                                    <button onclick="deleteSubUnit(${subject.id}, ${u.id}, ${sub.id})" style="background: none; border: none; color: #ef4444; cursor: pointer; font-size: 11px;">[ลบ]</button>
                                </li>
                            `).join('') +
                        `</ul>`;
                    } else {
                        subHtml = `<div style="font-size: 12px; color: #9ca3af; margin-left: 20px;">ยังไม่มีหน่วยย่อย</div>`;
                    }

                    return `
                        <div style="background: #ffffff; border: 1px dashed #cbd5e1; padding: 10px; border-radius: 6px; margin-bottom: 8px;">
                            <div style="display: flex; justify-content: space-between; align-items: center;">
                                <span><b>${u.name}</b> <span style="color: #2563eb;">(คะแนนบทรวม: ${u.maxScore} คะแนน)</span></span>
                                <div>
                                    <button onclick="addSubUnit(${subject.id}, ${u.id})" style="background: #10b981; color: white; border: none; padding: 3px 8px; border-radius: 4px; font-size: 12px; cursor: pointer;">+ เพิ่มหน่วยย่อย</button>
                                    <button onclick="deleteUnit(${subject.id}, ${u.id})" style="background: #ef4444; color: white; border: none; padding: 3px 8px; border-radius: 4px; font-size: 12px; cursor: pointer; margin-left: 5px;">ลบบท</button>
                                </div>
                            </div>
                            ${subHtml}
                        </div>
                    `;
                }).join('') +
            `</div>`;
        } else {
            unitsHtml = `<p style="color: #9ca3af; font-size: 14px; margin: 5px 0 0 0;">ยังไม่มีบทเรียน (โควต้าคะแนนเก็บรวม ${totalUnitQuota} คะแนน)</p>`;
        }

        const gradeBadge = subject.grade ? `<span style="background: #2563eb; color: white; padding: 2px 8px; border-radius: 12px; font-size: 12px; margin-right: 6px;">${subject.grade}</span>` : "";

        card.innerHTML = `
            <div style="display: flex; justify-content: space-between; align-items: center;">
                <h3 style="margin: 0; color: #1d4ed8;">${gradeBadge}${subject.code} - ${subject.name}</h3>
                <button onclick="deleteSubject(${subject.id})" style="background: #ef4444; color: white; border: none; padding: 6px 12px; border-radius: 6px; cursor: pointer;">ลบวิชานี้</button>
            </div>
            <div style="margin-top: 8px; font-size: 14px;">
                <span>🎯 กลางภาค: <b>${subject.midtermMax || 0}</b> คะแนน | </span>
                <span>🏁 ปลายภาค: <b>${subject.finalMax || 0}</b> คะแนน | </span>
                <span>📖 โควต้าคะแนนเก็บรวม: <b>${totalUnitQuota}</b> คะแนน</span>
            </div>
            ${unitsHtml}
        `;

        container.appendChild(card);
    });
}

// อัปเดต Dropdown เลือกวิชาให้แสดงระดับชั้น
function renderSubjectSelectDropdown() {
    const select = document.getElementById("scoreSubjectSelect");
    const unitSelect = document.getElementById("unitSubjectSelect");

    const optionsHtml = `<option value="">-- เลือกวิชา --</option>` +
        subjects.map(s => `<option value="${s.id}">[${s.grade || 'ทั่วไป'}] ${s.code} - ${s.name}</option>`).join('');

    if (select) select.innerHTML = optionsHtml;
    if (unitSelect) unitSelect.innerHTML = optionsHtml;
}

function addUnit() {
    const subjectId = Number(document.getElementById("unitSubjectSelect").value);
    const unitName = document.getElementById("unitName").value.trim();

    if (!subjectId || !unitName) {
        alert("กรุณาเลือกวิชาและกรอกชื่อบทเรียน");
        return;
    }

    const subject = subjects.find(s => s.id === subjectId);
    if (subject) {
        if (!subject.units) subject.units = [];

        subject.units.push({
            id: Date.now(),
            name: unitName,
            subUnits: []
        });

        recalculateUnitScores(subject);
        saveData();
        document.getElementById("unitName").value = "";
        updateAll();
    }
}

function addSubUnit(subjectId, unitId) {
    const subName = prompt("กรุณากรอกชื่อหน่วยย่อย (เช่น ใบงานที่ 1, ทดสอบย่อย):");
    if (!subName) return;

    const subject = subjects.find(s => s.id === subjectId);
    if (subject) {
        const unit = subject.units.find(u => u.id === unitId);
        if (unit) {
            if (!unit.subUnits) unit.subUnits = [];
            unit.subUnits.push({
                id: Date.now(),
                name: subName
            });

            recalculateSubUnitScores(unit);
            saveData();
            updateAll();
        }
    }
}

function recalculateUnitScores(subject) {
    if (!subject.units || subject.units.length === 0) return;

    const remainingScore = 100 - ((subject.midtermMax || 0) + (subject.finalMax || 0));
    if (remainingScore <= 0) {
        subject.units.forEach(unit => unit.maxScore = 0);
        return;
    }

    const unitCount = subject.units.length;
    const baseScore = Math.floor((remainingScore / unitCount) * 100) / 100;
    let sum = baseScore * unitCount;
    let diff = Math.round((remainingScore - sum) * 100) / 100;

    subject.units.forEach((unit, index) => {
        if (index === 0 && diff > 0) {
            unit.maxScore = Number((baseScore + diff).toFixed(2));
        } else {
            unit.maxScore = baseScore;
        }
        recalculateSubUnitScores(unit);
    });
}

function recalculateSubUnitScores(unit) {
    if (!unit.subUnits || unit.subUnits.length === 0) return;

    const subCount = unit.subUnits.length;
    const baseSubScore = Math.floor((unit.maxScore / subCount) * 100) / 100;
    let sum = baseSubScore * subCount;
    let diff = Math.round((unit.maxScore - sum) * 100) / 100;

    unit.subUnits.forEach((sub, index) => {
        if (index === 0 && diff > 0) {
            sub.maxScore = Number((baseSubScore + diff).toFixed(2));
        } else {
            sub.maxScore = baseSubScore;
        }
    });
}

function deleteUnit(subjectId, unitId) {
    if (!confirm("ต้องการลบบทเรียนนี้หรือไม่?")) return;

    const subject = subjects.find(s => s.id === subjectId);
    if (subject) {
        subject.units = subject.units.filter(u => u.id !== unitId);
        scores = scores.filter(s => !(s.subjectId === subjectId && s.unitId === unitId));
        
        recalculateUnitScores(subject);
        saveData();
        updateAll();
    }
}

function deleteSubUnit(subjectId, unitId, subUnitId) {
    if (!confirm("ต้องการลบหน่วยย่อยนี้หรือไม่?")) return;

    const subject = subjects.find(s => s.id === subjectId);
    if (subject) {
        const unit = subject.units.find(u => u.id === unitId);
        if (unit) {
            unit.subUnits = unit.subUnits.filter(sub => sub.id !== subUnitId);
            scores = scores.filter(s => !(s.subjectId === subjectId && s.unitId === unitId && s.subUnitId === subUnitId));
            
            recalculateSubUnitScores(unit);
            saveData();
            updateAll();
        }
    }
}

function deleteSubject(id) {
    if (!confirm("ต้องการลบวิชานี้หรือไม่?")) return;
    subjects = subjects.filter(subject => subject.id !== id);
    scores = scores.filter(score => score.subjectId !== id);
    saveData();
    updateAll();
}

function renderSubjects() {
    const container = document.getElementById("subjectListContainer");
    const dropdown = document.getElementById("unitSubjectSelect");

    if (dropdown) dropdown.innerHTML = `<option value="">-- เลือกวิชา --</option>`;
    if (!container) return;

    container.innerHTML = "";
    if (subjects.length === 0) {
        container.innerHTML = "<p style='color: #888;'>ยังไม่มีข้อมูลวิชา</p>";
        return;
    }

    subjects.forEach(subject => {
        if (dropdown) {
            const option = document.createElement("option");
            option.value = subject.id;
            option.textContent = `${subject.code} - ${subject.name}`;
            dropdown.appendChild(option);
        }

        const totalUnitQuota = 100 - ((subject.midtermMax || 0) + (subject.finalMax || 0));

        const card = document.createElement("div");
        card.style.cssText = "background: #f9fafb; border: 1px solid #e5e7eb; border-radius: 8px; padding: 15px; margin-bottom: 15px;";
        
        let unitsHtml = "";
        if (subject.units && subject.units.length > 0) {
            unitsHtml = `<div style="margin-top: 10px;">` +
                subject.units.map(u => {
                    let subHtml = "";

                    if (u.subUnits && u.subUnits.length > 0) {
                        subHtml = `<ul style="margin: 5px 0 10px 20px; font-size: 13px; color: #4b5563;">` +
                            u.subUnits.map(sub => `
                                <li>
                                    ${sub.name} <span style="color: #059669; font-weight: bold;">(คำนวณเต็ม: ${sub.maxScore} คะแนน)</span>
                                    <button onclick="deleteSubUnit(${subject.id}, ${u.id}, ${sub.id})" style="background: none; border: none; color: #ef4444; cursor: pointer; font-size: 11px;">[ลบ]</button>
                                </li>
                            `).join('') +
                        `</ul>`;
                    } else {
                        subHtml = `<div style="font-size: 12px; color: #9ca3af; margin-left: 20px;">ยังไม่มีหน่วยย่อย</div>`;
                    }

                    return `
                        <div style="background: #ffffff; border: 1px dashed #cbd5e1; padding: 10px; border-radius: 6px; margin-bottom: 8px;">
                            <div style="display: flex; justify-content: space-between; align-items: center;">
                                <span><b>${u.name}</b> <span style="color: #2563eb;">(คะแนนบทรวม: ${u.maxScore} คะแนน)</span></span>
                                <div>
                                    <button onclick="addSubUnit(${subject.id}, ${u.id})" style="background: #10b981; color: white; border: none; padding: 3px 8px; border-radius: 4px; font-size: 12px; cursor: pointer;">+ เพิ่มหน่วยย่อย</button>
                                    <button onclick="deleteUnit(${subject.id}, ${u.id})" style="background: #ef4444; color: white; border: none; padding: 3px 8px; border-radius: 4px; font-size: 12px; cursor: pointer; margin-left: 5px;">ลบบท</button>
                                </div>
                            </div>
                            ${subHtml}
                        </div>
                    `;
                }).join('') +
            `</div>`;
        } else {
            unitsHtml = `<p style="color: #9ca3af; font-size: 14px; margin: 5px 0 0 0;">ยังไม่มีบทเรียน (โควต้าคะแนนเก็บรวม ${totalUnitQuota} คะแนน)</p>`;
        }

        card.innerHTML = `
            <div style="display: flex; justify-content: space-between; align-items: center;">
                <h3 style="margin: 0; color: #1d4ed8;">${subject.code} - ${subject.name}</h3>
                <button onclick="deleteSubject(${subject.id})" style="background: #ef4444; color: white; border: none; padding: 6px 12px; border-radius: 6px; cursor: pointer;">ลบวิชานี้</button>
            </div>
            <div style="margin-top: 8px; font-size: 14px;">
                <span>🎯 กลางภาค: <b>${subject.midtermMax || 0}</b> คะแนน | </span>
                <span>🏁 ปลายภาค: <b>${subject.finalMax || 0}</b> คะแนน | </span>
                <span>📖 โควต้าคะแนนเก็บรวม: <b>${totalUnitQuota}</b> คะแนน</span>
            </div>
            ${unitsHtml}
        `;

        container.appendChild(card);
    });
}


// ==============================
// การสร้างและบันทึกตาราง Matrix
// ==============================

function renderSubjectSelectDropdown() {
    const select = document.getElementById("scoreSubjectSelect");
    if (!select) return;

    select.innerHTML = `<option value="">-- เลือกวิชา --</option>`;
    subjects.forEach(subject => {
        const option = document.createElement("option");
        option.value = subject.id;
        option.textContent = `${subject.code} - ${subject.name}`;
        select.appendChild(option);
    });
}

// ตัวอย่างการปรับฟังก์ชันเรนเดอร์ช่องคะแนนในตาราง Matrix ให้เป็น Checkbox
// 1. ฟังก์ชันเรนเดอร์ตาราง Matrix (เพิ่มช่องกรอกกลางภาค/ปลายภาค)
function renderScoreMatrix() {
    const container = document.getElementById('matrixContainer');
    const selectedSubjectId = document.getElementById('scoreSubjectSelect').value;

    if (!selectedSubjectId) {
        container.innerHTML = '<p style="color: #6b7280; text-align: center; padding: 20px;">กรุณาเลือกวิชาด้านบนเพื่อเริ่มกรอกคะแนน</p>';
        return;
    }

    const subject = subjects.find(s => s.id == selectedSubjectId);
    if (!subject || !subject.units || subject.units.length === 0) {
        container.innerHTML = '<p style="color: #6b7280; text-align: center; padding: 20px;">ไม่พบข้อมูลบทเรียนในวิชานี้</p>';
        return;
    }

    // กรองรายชื่อนักเรียนตามระดับชั้นวิชา (ถ้ามีกำหนด)
    let filteredStudents = students;
    if (subject.grade) {
        filteredStudents = students.filter(student => {
            if (!student.className) return true;
            return student.className.startsWith(subject.grade);
        });
    }

    if (filteredStudents.length === 0) {
        container.innerHTML = `<p style="color: #ef4444; text-align: center; padding: 20px;">ไม่พบรายชื่อนักเรียนระดับชั้น ${subject.grade || ''} ในระบบ</p>`;
        return;
    }

    // รวบรวมหน่วยย่อยทั้งหมด
    let unitGroups = [];
    let allSubUnits = [];

    subject.units.forEach(unit => {
        let currentSubUnits = [];
        if (unit.subUnits && unit.subUnits.length > 0) {
            unit.subUnits.forEach(sub => {
                const subObj = {
                    unitId: unit.id,
                    unitName: unit.name,
                    subUnitId: sub.id,
                    subUnitName: sub.name,
                    maxScore: sub.maxScore || sub.score || (unit.maxScore ? (unit.maxScore / unit.subUnits.length).toFixed(2) : 2)
                };
                currentSubUnits.push(subObj);
                allSubUnits.push(subObj);
            });
        } else {
            const subObj = {
                unitId: unit.id,
                unitName: unit.name,
                subUnitId: unit.id,
                subUnitName: unit.name,
                maxScore: unit.maxScore || unit.score || 10
            };
            currentSubUnits.push(subObj);
            allSubUnits.push(subObj);
        }

        unitGroups.push({
            unitId: unit.id,
            unitName: unit.name,
            colSpan: currentSubUnits.length,
            subUnits: currentSubUnits
        });
    });

    const midtermMax = subject.midtermScore || 20;
    const finalMax = subject.finalScore || 30;

    // สร้างโครงสร้างหัวตาราง (เพิ่มคอลัมน์ กลางภาค และ ปลายภาค)
    let tableHtml = `<table border="1" class="score-matrix-table" style="width: 100%; border-collapse: collapse; text-align: center; font-size: 13px;">
        <thead>
            <tr style="background: #e2e8f0; color: #1e293b;">
                <th rowspan="2" style="padding: 10px; min-width: 160px; position: sticky; left: 0; background: #e2e8f0; z-index: 3; border-right: 2px solid #cbd5e1;">ชื่อ-นามสกุล</th>`;

    unitGroups.forEach(group => {
        tableHtml += `<th colspan="${group.colSpan}" style="padding: 8px; font-weight: bold; font-size: 13px; background: #dbeafe; color: #1e40af; border-left: 2px solid #93c5fd; border-right: 2px solid #93c5fd;">
            📖 ${group.unitName}
        </th>`;
    });

    // เพิ่มหัวข้อกลางภาค/ปลายภาคแถวบน
    tableHtml += `<th colspan="2" style="padding: 8px; font-weight: bold; font-size: 13px; background: #fef08a; color: #854d0e; border-left: 2px solid #fde047;">🎯 คะแนนสอบ</th>
    </tr>
    <tr style="background: #f8fafc; color: #475569;">`;

    allSubUnits.forEach(sub => {
        tableHtml += `<th style="padding: 6px; font-size: 11px; font-weight: normal; min-width: 110px; border-bottom: 2px solid #cbd5e1;">
            <strong>${sub.subUnitName}</strong><br>
            <span style="color: #2563eb; font-size: 10px;">(เต็ม ${sub.maxScore})</span>
        </th>`;
    });

    // เพิ่มหัวข้อช่องกรอกกลางภาค/ปลายภาคแถวล่าง
    tableHtml += `<th style="padding: 6px; min-width: 90px; background: #fef9c3; border-bottom: 2px solid #cbd5e1;">กลางภาค<br><span style="color:#854d0e; font-size:10px;">(เต็ม ${midtermMax})</span></th>
                  <th style="padding: 6px; min-width: 90px; background: #fef9c3; border-bottom: 2px solid #cbd5e1;">ปลายภาค<br><span style="color:#854d0e; font-size:10px;">(เต็ม ${finalMax})</span></th>
    </tr></thead><tbody>`;

    // วนลูปสร้างแถวนักเรียน
    filteredStudents.forEach((student, idx) => {
        const rowBg = idx % 2 === 0 ? "#ffffff" : "#f9fafb";
        tableHtml += `<tr style="background: ${rowBg};">
            <td style="padding: 8px; text-align: left; font-weight: bold; position: sticky; left: 0; background: ${rowBg}; z-index: 1; border-right: 2px solid #cbd5e1;">${student.name} <span style="font-size: 11px; color: #64748b; font-weight: normal;">(${student.className || ''})</span></td>`;

        // 1. ช่องติ๊กเก็บคะแนนเก็บหน่วยย่อย
        allSubUnits.forEach(sub => {
            const currentRecord = scores.find(s => 
                s.studentId == student.id && 
                s.subjectId == subject.id && 
                (s.subUnitId == sub.subUnitId || s.unitKey == `${sub.unitId}_${sub.subUnitId}` || s.unitId == sub.subUnitId)
            );
            
            const isChecked = currentRecord && Number(currentRecord.score) > 0 ? "checked" : "";

            tableHtml += `<td style="padding: 8px;">
                <input type="checkbox" 
                    ${isChecked} 
                    style="transform: scale(1.4); cursor: pointer;"
                    onchange="toggleSubUnitScore(this, '${student.id}', '${subject.id}', '${sub.unitId}', '${sub.subUnitId}', ${sub.maxScore})">
            </td>`;
        });

        // 2. ดึงคะแนนกลางภาค / ปลายภาคเดิม
        const midRecord = scores.find(s => s.studentId == student.id && s.subjectId == subject.id && s.unitId == 'midterm');
        const finalRecord = scores.find(s => s.studentId == student.id && s.subjectId == subject.id && s.unitId == 'final');

        const midVal = midRecord ? midRecord.score : "";
        const finalVal = finalRecord ? finalRecord.score : "";

        // ช่องกรอกคะแนนสอบกลางภาค และปลายภาค
        tableHtml += `<td style="padding: 4px; background: #fffbebfb;">
            <input type="number" value="${midVal}" min="0" max="${midtermMax}" placeholder="0" style="width: 60px; text-align: center; padding: 4px;"
            onchange="saveExamScore('${student.id}', '${subject.id}', 'midterm', this.value, ${midtermMax})">
        </td>
        <td style="padding: 4px; background: #fffbebfb;">
            <input type="number" value="${finalVal}" min="0" max="${finalMax}" placeholder="0" style="width: 60px; text-align: center; padding: 4px;"
            onchange="saveExamScore('${student.id}', '${subject.id}', 'final', this.value, ${finalMax})">
        </td>`;

        tableHtml += `</tr>`;
    });

    tableHtml += `</tbody></table>`;
    container.innerHTML = tableHtml;

    // คำนวณสรุปเกรดทันทีที่เรนเดอร์ตาราง
    calculateAndRenderSummaryScores();
}

// 2. ฟังก์ชันอัปเดตคะแนนจากการติ๊กถูก + สั่งให้คำนวณเกรดทันที
async function toggleSubUnitScore(checkbox, studentId, subjectId, unitId, subUnitId, maxScore) {
    const finalScore = checkbox.checked ? Number(maxScore) : 0;
    const unitKey = `${unitId}_${subUnitId}`;

    let existingIndex = scores.findIndex(s => 
        s.studentId == studentId && 
        s.subjectId == subjectId && 
        (s.subUnitId == subUnitId || s.unitKey == unitKey)
    );

    if (existingIndex >= 0) {
        scores[existingIndex].score = finalScore;
    } else {
        scores.push({
            studentId: studentId,
            subjectId: subjectId,
            unitId: unitId,
            subUnitId: subUnitId,
            unitKey: unitKey,
            score: finalScore
        });
    }

    // คำนวณคะแนนรวมและเกรดใหม่ทันที
    calculateAndRenderSummaryScores();

    // บันทึกลง Google Sheets
    try {
        await fetch(GAS_API_URL, {
            method: "POST",
            body: JSON.stringify({
                action: "saveScore",
                studentId: studentId,
                subjectId: subjectId,
                unitKey: unitKey,
                unitId: unitId,
                subUnitId: subUnitId,
                score: finalScore
            })
        });
    } catch (e) {
        console.error("บันทึกคะแนนล้มเหลว:", e);
    }
}

// 3. ฟังก์ชันบันทึกคะแนนสอบกลางภาค / ปลายภาค
async function saveExamScore(studentId, subjectId, examType, value, maxScore) {
    let scoreVal = Number(value);
    if (isNaN(scoreVal) || scoreVal < 0) scoreVal = 0;
    if (scoreVal > maxScore) {
        alert(`คะแนนต้องไม่เกิน ${maxScore}`);
        scoreVal = maxScore;
    }

    let existingIndex = scores.findIndex(s => s.studentId == studentId && s.subjectId == subjectId && s.unitId == examType);
    if (existingIndex >= 0) {
        scores[existingIndex].score = scoreVal;
    } else {
        scores.push({
            studentId: studentId,
            subjectId: subjectId,
            unitId: examType,
            unitKey: examType,
            score: scoreVal
        });
    }

    // คำนวณคะแนนรวมและเกรดใหม่ทันที
    calculateAndRenderSummaryScores();

    // บันทึกลง Google Sheets
    try {
        await fetch(GAS_API_URL, {
            method: "POST",
            body: JSON.stringify({
                action: "saveScore",
                studentId: studentId,
                subjectId: subjectId,
                unitKey: examType,
                unitId: examType,
                score: scoreVal
            })
        });
    } catch (e) {
        console.error("บันทึกคะแนนสอบล้มเหลว:", e);
    }
}

// 4. ฟังก์ชันคำนวณสรุปคะแนนรวม เกรด และรายงานแสดงผลด้านล่าง
// ==========================================
// ฟังก์ชันเรนเดอร์ตารางสรุปคะแนนรวมและเกรด (คำนวณสัดส่วนโควตาจริง)
// ==========================================
function calculateAndRenderSummaryScores() {
    const selectedSubjectId = document.getElementById('scoreSubjectSelect')?.value;
    if (!selectedSubjectId) return;

    const subject = subjects.find(s => String(s.id).trim() === String(selectedSubjectId).trim());
    if (!subject) return;

    const scoreTable = document.getElementById('scoreTable');
    if (!scoreTable) return;

    scoreTable.innerHTML = "";

    // 1. ดึงคะแนนเต็มกลางภาค และ ปลายภาค
    const midMax = Number(subject.midtermMax || subject.midtermScore || 20);
    const finalMax = Number(subject.finalMax || subject.finalScore || 20);

    // 2. คำนวณโควตาคะแนนเก็บที่เหลือ (เช่น 100 - 20 - 20 = 60)
    const remainingQuota = Math.max(0, 100 - midMax - finalMax);

    // 3. หาผลรวมคะแนนเต็มดิบของทุกหน่วยย่อยในวิชานี้
    let totalRawUnitsMax = 0;
    if (subject.units && subject.units.length > 0) {
        subject.units.forEach(unit => {
            if (unit.subUnits && unit.subUnits.length > 0) {
                unit.subUnits.forEach(sub => {
                    totalRawUnitsMax += Number(sub.maxScore || sub.score || 10);
                });
            } else {
                totalRawUnitsMax += Number(unit.maxScore || unit.score || 10);
            }
        });
    }

    let filteredStudents = students;
    if (subject.grade) {
        filteredStudents = students.filter(s => !s.className || s.className.startsWith(subject.grade));
    }

    filteredStudents.forEach(student => {
        // 4. คำนวณคะแนนดิบที่ติ๊กได้ในส่วนบทเรียน
        let studentRawUnitsScore = 0;
        if (subject.units && subject.units.length > 0) {
            subject.units.forEach(unit => {
                if (unit.subUnits && unit.subUnits.length > 0) {
                    unit.subUnits.forEach(sub => {
                        const rec = scores.find(s => 
                            String(s.studentId).trim() === String(student.id).trim() && 
                            String(s.subjectId).trim() === String(subject.id).trim() && 
                            (String(s.subUnitId).trim() === String(sub.id).trim() || s.unitKey === `${unit.id}_${sub.id}`)
                        );
                        if (rec && Number(rec.score) > 0) {
                            // ใช้ค่า rawScore ถ้ามี หรือใช้ maxScore ของ subUnit เมื่อติ๊กถูก
                            studentRawUnitsScore += Number(rec.rawScore || sub.maxScore || sub.score || 10);
                        }
                    });
                } else {
                    const rec = scores.find(s => 
                        String(s.studentId).trim() === String(student.id).trim() && 
                        String(s.subjectId).trim() === String(subject.id).trim() && 
                        String(s.unitId).trim() === String(unit.id).trim() && !s.subUnitId
                    );
                    if (rec && Number(rec.score) > 0) {
                        studentRawUnitsScore += Number(rec.rawScore || unit.maxScore || unit.score || 10);
                    }
                }
            });
        }

        // 5. ทอนสัดส่วนคะแนนเก็บให้อยู่ในโควตาที่เหลือ (เช่น ได้เต็มดิบ -> ทอนเหลือ 60)
        let weightedUnitsScore = totalRawUnitsMax > 0 
            ? (studentRawUnitsScore / totalRawUnitsMax) * remainingQuota 
            : 0;

        // 6. ดึงคะแนนสอบกลางภาค / ปลายภาค จริง
        const midRec = scores.find(s => 
            String(s.studentId).trim() === String(student.id).trim() && 
            String(s.subjectId).trim() === String(subject.id).trim() && 
            (s.unitValue === 'midterm' || s.unitId === 'midterm' || s.unitKey === 'midterm')
        );
        const finalRec = scores.find(s => 
            String(s.studentId).trim() === String(student.id).trim() && 
            String(s.subjectId).trim() === String(subject.id).trim() && 
            (s.unitValue === 'final' || s.unitId === 'final' || s.unitKey === 'final')
        );

        const midScore = midRec ? Number(midRec.score || 0) : 0;
        const finalScore = finalRec ? Number(finalRec.score || 0) : 0;

        // 7. คะแนนรวมสุทธิทั้งหมด
        const finalTotalScore = weightedUnitsScore + midScore + finalScore;

        // 8. คำนวณเกรด
        let grade = "0";
        if (finalTotalScore >= 80) grade = "4";
        else if (finalTotalScore >= 75) grade = "3.5";
        else if (finalTotalScore >= 70) grade = "3";
        else if (finalTotalScore >= 65) grade = "2.5";
        else if (finalTotalScore >= 60) grade = "2";
        else if (finalTotalScore >= 55) grade = "1.5";
        else if (finalTotalScore >= 50) grade = "1";

        scoreTable.innerHTML += `<tr>
            <td style="padding: 8px;">${student.code || student.id}</td>
            <td style="padding: 8px; font-weight: bold;">${student.name}</td>
            <td style="padding: 8px;">${subject.name}</td>
            <td style="padding: 8px; font-weight: bold; color: #2563eb;">${finalTotalScore.toFixed(2)} / 100</td>
            <td style="padding: 8px;">${finalTotalScore.toFixed(2)}%</td>
            <td style="padding: 8px; font-weight: bold; color: ${grade === '0' ? '#ef4444' : '#10b981'};">${grade}</td>
        </tr>`;
    });
}

// ==========================================
// 1. ฟังก์ชันบันทึกการติ๊ก Checkbox (ทอนสัดส่วนตามโควตาคะแนนเก็บทันที)
// ==========================================
async function toggleSubUnitScore(checkbox, studentId, subjectId, unitId, subUnitId, rawMaxScore) {
    const subject = subjects.find(s => String(s.id).trim() === String(subjectId).trim());
    if (!subject) return;

    // คำนวณโควตาคะแนนเก็บที่เหลือ (เช่น 100 - กลางภาค 20 - ปลายภาค 20 = 60 คะแนน)
    const midMax = Number(subject.midtermMax || subject.midtermScore || 20);
    const finalMax = Number(subject.finalMax || subject.finalScore || 20);
    const remainingQuota = Math.max(0, 100 - midMax - finalMax);

    // หาคะแนนดิบรวมทั้งหมดของวิชานี้
    let totalRawUnitsMax = 0;
    if (subject.units && subject.units.length > 0) {
        subject.units.forEach(u => {
            if (u.subUnits && u.subUnits.length > 0) {
                u.subUnits.forEach(s => { totalRawUnitsMax += Number(s.maxScore || s.score || 10); });
            } else {
                totalRawUnitsMax += Number(u.maxScore || u.score || 10);
            }
        });
    }

    // คำนวณน้ำหนักคะแนนต่อน้ำหนักดิบ 1 หน่วย
    const weightFactor = totalRawUnitsMax > 0 ? (remainingQuota / totalRawUnitsMax) : 1;

    // คะแนนดิบช่องนี้ x น้ำหนักสัดส่วนจริง (เช่น ติ๊กได้ดิบ 10 -> บันทึกจริงเป็นสัดส่วนของ 60)
    const rawScoreVal = checkbox.checked ? Number(rawMaxScore) : 0;
    const finalWeightedScore = rawScoreVal * weightFactor;

    const unitKey = `${unitId}_${subUnitId}`;

    let existingIndex = scores.findIndex(s => 
        String(s.studentId).trim() === String(studentId).trim() && 
        String(s.subjectId).trim() === String(subjectId).trim() && 
        (String(s.subUnitId).trim() === String(subUnitId).trim() || s.unitKey === unitKey)
    );

    if (existingIndex >= 0) {
        scores[existingIndex].score = finalWeightedScore;
        scores[existingIndex].rawScore = rawScoreVal; // เก็บค่าดิบไว้เช็กติ๊ก
    } else {
        scores.push({
            studentId: studentId,
            subjectId: subjectId,
            unitId: unitId,
            subUnitId: subUnitId,
            unitKey: unitKey,
            score: finalWeightedScore,
            rawScore: rawScoreVal
        });
    }

    // คำนวณและอัปเดตตารางสรุปคะแนนสดๆ ทันที
    if (typeof calculateAndRenderSummaryScores === 'function') {
        calculateAndRenderSummaryScores();
    }

    // บันทึกลง Google Sheets (ถ้ามี)
    if (typeof GAS_API_URL !== 'undefined') {
        try {
            await fetch(GAS_API_URL, {
                method: "POST",
                body: JSON.stringify({
                    action: "saveScore",
                    studentId: studentId,
                    subjectId: subjectId,
                    unitKey: unitKey,
                    unitId: unitId,
                    subUnitId: subUnitId,
                    score: finalWeightedScore
                })
            });
        } catch (e) {
            console.error("บันทึกคะแนนล้มเหลว:", e);
        }
    }
}

// ==========================================
// 2. ฟังก์ชันเรนเดอร์ตาราง Matrix (แสดงคะแนนเต็มช่องละเท่าไหร่ตามสัดส่วน)
// ==========================================
function renderScoreMatrix() {
    const container = document.getElementById('matrixContainer');
    const selectedSubjectId = document.getElementById('scoreSubjectSelect')?.value;

    if (!selectedSubjectId) {
        container.innerHTML = '<p style="color: #6b7280; text-align: center; padding: 20px;">กรุณาเลือกวิชาด้านบนเพื่อเริ่มกรอกคะแนน</p>';
        return;
    }

    const subject = subjects.find(s => String(s.id).trim() === String(selectedSubjectId).trim());
    if (!subject || !subject.units || subject.units.length === 0) {
        container.innerHTML = '<p style="color: #6b7280; text-align: center; padding: 20px;">ไม่พบข้อมูลบทเรียนในวิชานี้</p>';
        return;
    }

    // คำนวณโควตาคะแนนเก็บและสัดส่วน
    const midMax = Number(subject.midtermMax || subject.midtermScore || 20);
    const finalMax = Number(subject.finalMax || subject.finalScore || 20);
    const remainingQuota = Math.max(0, 100 - midMax - finalMax);

    let totalRawUnitsMax = 0;
    subject.units.forEach(u => {
        if (u.subUnits && u.subUnits.length > 0) {
            u.subUnits.forEach(s => { totalRawUnitsMax += Number(s.maxScore || s.score || 10); });
        } else {
            totalRawUnitsMax += Number(u.maxScore || u.score || 10);
        }
    });

    const weightFactor = totalRawUnitsMax > 0 ? (remainingQuota / totalRawUnitsMax) : 1;

    let filteredStudents = students;
    if (subject.grade) {
        filteredStudents = students.filter(s => !s.className || s.className.startsWith(subject.grade));
    }

    let unitGroups = [];
    let allSubUnits = [];

    subject.units.forEach(unit => {
        let currentSubUnits = [];
        if (unit.subUnits && unit.subUnits.length > 0) {
            unit.subUnits.forEach(sub => {
                const rawMax = Number(sub.maxScore || sub.score || 10);
                const weightedMax = rawMax * weightFactor; // คะแนนเต็มช่องนี้แบบสัดส่วนจริง
                const subObj = {
                    unitId: unit.id,
                    unitName: unit.name,
                    subUnitId: sub.id,
                    subUnitName: sub.name,
                    rawMaxScore: rawMax,
                    maxScore: weightedMax.toFixed(2)
                };
                currentSubUnits.push(subObj);
                allSubUnits.push(subObj);
            });
        } else {
            const rawMax = Number(unit.maxScore || unit.score || 10);
            const weightedMax = rawMax * weightFactor;
            const subObj = {
                unitId: unit.id,
                unitName: unit.name,
                subUnitId: unit.id,
                subUnitName: unit.name,
                rawMaxScore: rawMax,
                maxScore: weightedMax.toFixed(2)
            };
            currentSubUnits.push(subObj);
            allSubUnits.push(subObj);
        }

        unitGroups.push({
            unitId: unit.id,
            unitName: unit.name,
            colSpan: currentSubUnits.length,
            subUnits: currentSubUnits
        });
    });

    let tableHtml = `<table border="1" class="score-matrix-table" style="width: 100%; border-collapse: collapse; text-align: center; font-size: 13px;">
        <thead>
            <tr style="background: #e2e8f0; color: #1e293b;">
                <th rowspan="2" style="padding: 10px; min-width: 160px; position: sticky; left: 0; background: #e2e8f0; z-index: 3; border-right: 2px solid #cbd5e1;">ชื่อ-นามสกุล</th>`;

    unitGroups.forEach(group => {
        tableHtml += `<th colspan="${group.colSpan}" style="padding: 8px; font-weight: bold; background: #dbeafe; color: #1e40af; border: 1px solid #93c5fd;">
            📖 ${group.unitName}
        </th>`;
    });

    tableHtml += `<th colspan="2" style="padding: 8px; font-weight: bold; background: #fef08a; color: #854d0e; border: 1px solid #fde047;">🎯 คะแนนสอบ</th>
    </tr>
    <tr style="background: #f8fafc; color: #475569;">`;

    allSubUnits.forEach(sub => {
        tableHtml += `<th style="padding: 6px; font-size: 11px; font-weight: normal; min-width: 100px; border-bottom: 2px solid #cbd5e1;">
            <strong>${sub.subUnitName}</strong><br>
            <span style="color: #2563eb; font-size: 10px;">(เต็ม ${sub.maxScore})</span>
        </th>`;
    });

    tableHtml += `<th style="padding: 6px; min-width: 85px; background: #fef9c3;">กลางภาค<br><span style="color:#854d0e; font-size:10px;">(เต็ม ${midMax})</span></th>
                  <th style="padding: 6px; min-width: 85px; background: #fef9c3;">ปลายภาค<br><span style="color:#854d0e; font-size:10px;">(เต็ม ${finalMax})</span></th>
    </tr></thead><tbody>`;

    filteredStudents.forEach((student, idx) => {
        const rowBg = idx % 2 === 0 ? "#ffffff" : "#f9fafb";
        tableHtml += `<tr style="background: ${rowBg};">
            <td style="padding: 8px; text-align: left; font-weight: bold; position: sticky; left: 0; background: ${rowBg}; z-index: 1; border-right: 2px solid #cbd5e1;">
                ${student.name} <span style="font-size: 11px; color: #64748b; font-weight: normal;">(${student.className || ''})</span>
            </td>`;

        allSubUnits.forEach(sub => {
            const currentRecord = scores.find(s => 
                String(s.studentId).trim() === String(student.id).trim() && 
                String(s.subjectId).trim() === String(subject.id).trim() && 
                (String(s.subUnitId).trim() === String(sub.subUnitId).trim() || s.unitKey === `${sub.unitId}_${sub.subUnitId}`)
            );
            
            const isChecked = currentRecord && Number(currentRecord.score) > 0 ? "checked" : "";

            tableHtml += `<td style="padding: 8px;">
                <input type="checkbox" 
                    ${isChecked} 
                    style="transform: scale(1.4); cursor: pointer;"
                    onchange="toggleSubUnitScore(this, '${student.id}', '${subject.id}', '${sub.unitId}', '${sub.subUnitId}', ${sub.rawMaxScore})">
            </td>`;
        });

        const midRecord = scores.find(s => String(s.studentId).trim() === String(student.id).trim() && String(s.subjectId).trim() === String(subject.id).trim() && (s.unitId === 'midterm' || s.unitKey === 'midterm'));
        const finalRecord = scores.find(s => String(s.studentId).trim() === String(student.id).trim() && String(s.subjectId).trim() === String(subject.id).trim() && (s.unitId === 'final' || s.unitKey === 'final'));

        tableHtml += `<td style="padding: 4px; background: #fffbebfb;">
            <input type="number" value="${midRecord ? midRecord.score : ''}" min="0" max="${midMax}" placeholder="0" style="width: 55px; text-align: center; padding: 4px;"
            onchange="saveExamScore('${student.id}', '${subject.id}', 'midterm', this.value, ${midMax})">
        </td>
        <td style="padding: 4px; background: #fffbebfb;">
            <input type="number" value="${finalRecord ? finalRecord.score : ''}" min="0" max="${finalMax}" placeholder="0" style="width: 55px; text-align: center; padding: 4px;"
            onchange="saveExamScore('${student.id}', '${subject.id}', 'final', this.value, ${finalMax})">
        </td>`;

        tableHtml += `</tr>`;
    });

    tableHtml += `</tbody></table>`;
    container.innerHTML = tableHtml;

    if (typeof calculateAndRenderSummaryScores === 'function') {
        calculateAndRenderSummaryScores();
    }
}

// ==========================================
// 3. ฟังก์ชันสรุปคะแนนรวมรวมทุกส่วน (คะแนนเก็บสัดส่วน + กลางภาค + ปลายภาค)
// ==========================================
function calculateAndRenderSummaryScores() {
    const selectedSubjectId = document.getElementById('scoreSubjectSelect')?.value;
    if (!selectedSubjectId) return;

    const subject = subjects.find(s => String(s.id).trim() === String(selectedSubjectId).trim());
    if (!subject) return;

    const scoreTable = document.getElementById('scoreTable');
    if (!scoreTable) return;

    scoreTable.innerHTML = "";

    let filteredStudents = students;
    if (subject.grade) {
        filteredStudents = students.filter(s => !s.className || s.className.startsWith(subject.grade));
    }

    filteredStudents.forEach(student => {
        // รวมคะแนนเก็บที่ผ่านการทอนสัดส่วนแล้ว
        const studentScores = scores.filter(s => String(s.studentId).trim() === String(student.id).trim() && String(s.subjectId).trim() === String(subject.id).trim());
        let totalScore = 0;

        studentScores.forEach(s => {
            totalScore += Number(s.score || 0);
        });

        let grade = "0";
        if (totalScore >= 80) grade = "4";
        else if (totalScore >= 75) grade = "3.5";
        else if (totalScore >= 70) grade = "3";
        else if (totalScore >= 65) grade = "2.5";
        else if (totalScore >= 60) grade = "2";
        else if (totalScore >= 55) grade = "1.5";
        else if (totalScore >= 50) grade = "1";

        scoreTable.innerHTML += `<tr>
            <td style="padding: 8px;">${student.code || student.id}</td>
            <td style="padding: 8px; font-weight: bold;">${student.name}</td>
            <td style="padding: 8px;">${subject.name}</td>
            <td style="padding: 8px; font-weight: bold; color: #2563eb;">${totalScore.toFixed(2)} / 100</td>
            <td style="padding: 8px;">${totalScore.toFixed(2)}%</td>
            <td style="padding: 8px; font-weight: bold; color: ${grade === '0' ? '#ef4444' : '#10b981'};">${grade}</td>
        </tr>`;
    });
}
function saveMatrixScore(studentId, subjectId, unitValue, unitId, subUnitId, maxAllowed, inputElem) {
    const scoreVal = inputElem.value.trim();

    if (scoreVal === "") {
        // ลบคะแนนออกถ้าช่องว่าง
        scores = scores.filter(s => !(s.studentId === studentId && s.subjectId === subjectId && s.unitValue === unitValue));
        saveData();
        renderMissingTasks();
        renderScores();
        return;
    }

    const numScore = Number(scoreVal);
    if (isNaN(numScore) || numScore < 0 || numScore > maxAllowed) {
        alert(`คะแนนต้องอยู่ระหว่าง 0 ถึง ${maxAllowed}`);
        inputElem.value = "";
        return;
    }

    const existingIndex = scores.findIndex(
        s => s.studentId === studentId && s.subjectId === subjectId && s.unitValue === unitValue
    );

    if (existingIndex !== -1) {
        scores[existingIndex].score = numScore;
    } else {
        scores.push({
            id: Date.now(),
            studentId: studentId,
            subjectId: subjectId,
            unitValue: unitValue,
            unitId: unitId,
            subUnitId: subUnitId,
            score: numScore
        });
    }

    saveData();
    renderMissingTasks();
    renderScores();
}


// ==============================
// คำนวณคะแนนรวมและเกรด
// ==============================

// ==============================
// คำนวณคะแนนรวมและเกรด (เพิ่มระบบเกรดตัวเลข 0 - 4)
// ==============================
// ==========================================
// ฟังก์ชันสร้างรายงาน PDF (สรุปคะแนนรวมแยกตามบทเรียน)
// ==========================================

// 1. ส่งออกรายงานรายคน
async function exportSingleStudentPDF() {
    const studentSelect = document.getElementById("singleStudentSelect") || document.getElementById("pdfStudentSelect");
    const subjectId = Number(document.getElementById("scoreSubjectSelect")?.value);

    if (!subjectId) {
        alert("กรุณาเลือกรายวิชาก่อนครับ");
        return;
    }
    const studentId = Number(studentSelect?.value);
    if (!studentId) {
        alert("กรุณาเลือกนักเรียนที่ต้องการออกรายงานครับ");
        return;
    }

    const student = students.find(s => s.id === studentId);
    const subject = subjects.find(s => s.id === subjectId);
    const printArea = document.getElementById("pdfPrintArea");

    printArea.style.display = "block";
    printArea.innerHTML = "";

    const card = buildStudentReportCard(student, subject);
    printArea.appendChild(card);

    const opt = {
        margin:       [8, 8, 8, 8],
        filename:     `รายงานคะแนน_${student.name}_${subject.code || subject.name}.pdf`,
        image:        { type: 'jpeg', quality: 0.98 },
        html2canvas:  { scale: 2, useCORS: true, scrollY: 0 },
        jsPDF:        { unit: 'mm', format: 'a4', orientation: 'portrait' },
        pagebreak:    { mode: ['avoid-all', 'css', 'legacy'] }
    };

    await html2pdf().set(opt).from(printArea).save();
    printArea.style.display = "none";
}

// 2. ส่งออกรายงานทั้งชั้นเรียน
async function exportClassPDF() {
    const subjectId = Number(document.getElementById("scoreSubjectSelect")?.value);

    if (!subjectId) {
        alert("กรุณาเลือกรายวิชาก่อนครับ");
        return;
    }

    const subject = subjects.find(s => s.id === subjectId);
    const filteredStudents = students.filter(student => {
        if (!subject.grade) return true;
        return student.className && student.className.startsWith(subject.grade);
    });

    if (filteredStudents.length === 0) {
        alert("ไม่พบนักเรียนในระดับชั้นนี้");
        return;
    }

    const printArea = document.getElementById("pdfPrintArea");
    printArea.style.display = "block";
    printArea.innerHTML = "";

    for (let i = 0; i < filteredStudents.length; i++) {
        const student = filteredStudents[i];
        const card = buildStudentReportCard(student, subject);
        if (i < filteredStudents.length - 1) {
            card.style.pageBreakAfter = "always";
        }
        printArea.appendChild(card);
    }

    const opt = {
        margin:       [8, 8, 8, 8],
        filename:     `รายงานคะแนนชั้นเรียน_${subject.grade || ''}_${subject.code || subject.name}.pdf`,
        image:        { type: 'jpeg', quality: 0.98 },
        html2canvas:  { scale: 2, useCORS: true, scrollY: 0 },
        jsPDF:        { unit: 'mm', format: 'a4', orientation: 'portrait' },
        pagebreak:    { mode: ['avoid-all', 'css', 'legacy'] }
    };

    await html2pdf().set(opt).from(printArea).save();
    printArea.style.display = "none";
}

// ==========================================
// ฟังก์ชันสร้างการ์ดรายงาน (ใช้ calculateSubjectTotal)
// ==========================================
// ==========================================
// ฟังก์ชันคำนวณคะแนนรวมและเกรดหลัก (ปรับสัดส่วนรวม 100 คะแนนเต็ม)
// ==========================================
function calculateSubjectTotal(studentId, subjectId) {
    const subject = subjects.find(s => String(s.id).trim() === String(subjectId).trim());
    if (!subject) return { totalScore: "0.00", gradeNum: "0", gradeLetter: "F" };

    // 1. ดึงคะแนนสอบกลางภาค / ปลายภาค
    const midMax = Number(subject.midtermMax || subject.midtermScore || 20);
    const finalMax = Number(subject.finalMax || subject.finalScore || 20);

    // 2. คำนวณโควตาคะแนนเก็บที่เหลือ (เช่น 100 - 20 - 20 = 60 คะแนน)
    const remainingQuota = Math.max(0, 100 - midMax - finalMax);

    // 3. คำนวณคะแนนดิบรวมทั้งหมดของบทเรียนที่มีในรายวิชานี้ (Raw Max Score Total)
    let totalRawUnitsMax = 0;
    if (subject.units && subject.units.length > 0) {
        subject.units.forEach(unit => {
            if (unit.subUnits && unit.subUnits.length > 0) {
                unit.subUnits.forEach(sub => {
                    totalRawUnitsMax += Number(sub.maxScore || sub.score || 10);
                });
            } else {
                totalRawUnitsMax += Number(unit.maxScore || unit.score || 10);
            }
        });
    }

    // 4. คำนวณคะแนนดิบที่นักเรียนทำได้ในส่วนบทเรียน
    let studentRawUnitsScore = 0;
    if (subject.units && subject.units.length > 0) {
        subject.units.forEach(unit => {
            if (unit.subUnits && unit.subUnits.length > 0) {
                unit.subUnits.forEach(sub => {
                    const rec = scores.find(s => 
                        String(s.studentId).trim() === String(studentId).trim() && 
                        String(s.subjectId).trim() === String(subjectId).trim() && 
                        (String(s.subUnitId).trim() === String(sub.id).trim() || s.unitKey === `${unit.id}_${sub.id}`)
                    );
                    if (rec) studentRawUnitsScore += Number(rec.score || 0);
                });
            } else {
                const rec = scores.find(s => 
                    String(s.studentId).trim() === String(studentId).trim() && 
                    String(s.subjectId).trim() === String(subjectId).trim() && 
                    String(s.unitId).trim() === String(unit.id).trim() && !s.subUnitId
                );
                if (rec) studentRawUnitsScore += Number(rec.score || 0);
            }
        });
    }

    // 5. ทอนสัดส่วนคะแนนเก็บให้อยู่ในโควตาที่เหลือ
    let weightedUnitsScore = totalRawUnitsMax > 0 
        ? (studentRawUnitsScore / totalRawUnitsMax) * remainingQuota 
        : 0;

    // 6. ดึงคะแนนสอบกลางภาคและปลายภาคจริง
    const midRec = scores.find(s => 
        String(s.studentId).trim() === String(studentId).trim() && 
        String(s.subjectId).trim() === String(subjectId).trim() && 
        (s.unitValue === 'midterm' || s.unitId === 'midterm' || s.unitKey === 'midterm')
    );
    const finalRec = scores.find(s => 
        String(s.studentId).trim() === String(studentId).trim() && 
        String(s.subjectId).trim() === String(subjectId).trim() && 
        (s.unitValue === 'final' || s.unitId === 'final' || s.unitKey === 'final')
    );

    const midScore = midRec ? Number(midRec.score || 0) : 0;
    const finalScore = finalRec ? Number(finalRec.score || 0) : 0;

    // 7. คำนวณคะแนนสุทธิรวมทั้งหมด (เต็ม 100)
    const finalTotalScore = weightedUnitsScore + midScore + finalScore;

    // 8. คำนวณระดับผลการเรียน (เกรด)
    let gradeNum = "0";
    let gradeLetter = "F";

    if (finalTotalScore >= 80) { gradeNum = "4"; gradeLetter = "A"; }
    else if (finalTotalScore >= 75) { gradeNum = "3.5"; gradeLetter = "B+"; }
    else if (finalTotalScore >= 70) { gradeNum = "3"; gradeLetter = "B"; }
    else if (finalTotalScore >= 65) { gradeNum = "2.5"; gradeLetter = "C+"; }
    else if (finalTotalScore >= 60) { gradeNum = "2"; gradeLetter = "C"; }
    else if (finalTotalScore >= 55) { gradeNum = "1.5"; gradeLetter = "D+"; }
    else if (finalTotalScore >= 50) { gradeNum = "1"; gradeLetter = "D"; }

    return {
        totalScore: finalTotalScore.toFixed(2),
        gradeNum: gradeNum,
        gradeLetter: gradeLetter
    };
}
// ==========================================
// ฟังก์ชันสร้างการ์ดรายงาน (ปรับสัดส่วนคะแนนเก็บให้รวมได้ 100 พอดี)
// ==========================================

function buildStudentReportCard(student, subject) {
    const cardContainer = document.createElement("div");
    cardContainer.style.cssText = "padding: 10px 15px; font-family: 'Sarabun', sans-serif; color: #1e293b; background: #fff; box-sizing: border-box;";

    // 1. ดึงคะแนนเต็มกลางภาคและปลายภาคที่กำหนดไว้
    const midMax = Number(subject.midtermMax || subject.midtermScore || 20);
    const finalMax = Number(subject.finalMax || subject.finalScore || 20);

    // 2. คำนวณโควตาคะแนนเก็บที่เหลือ (เช่น 100 - 20 - 20 = 60 คะแนน)
    const remainingQuota = Math.max(0, 100 - midMax - finalMax);

    // 3. คำนวณคะแนนดิบรวมของทุกบทเรียน (Raw Scores Sum)
    let totalRawUnitsMax = 0;
    if (subject.units && subject.units.length > 0) {
        subject.units.forEach(unit => {
            if (unit.subUnits && unit.subUnits.length > 0) {
                unit.subUnits.forEach(sub => {
                    totalRawUnitsMax += Number(sub.maxScore || sub.score || 10);
                });
            } else {
                totalRawUnitsMax += Number(unit.maxScore || unit.score || 10);
            }
        });
    }

    // 4. วนลูปคำนวณคะแนนแต่ละบทเรียนตามสัดส่วน
    let rowsHtml = "";
    let totalWeightedUnitScore = 0;

    if (subject.units && subject.units.length > 0) {
        subject.units.forEach(unit => {
            let rawUnitScore = 0;
            let rawUnitMax = 0;

            if (unit.subUnits && unit.subUnits.length > 0) {
                unit.subUnits.forEach(sub => {
                    const subMax = Number(sub.maxScore || sub.score || 10);
                    rawUnitMax += subMax;

                    const rec = scores.find(s => s.studentId === student.id && s.subjectId === subject.id && s.subUnitId === sub.id);
                    if (rec) rawUnitScore += Number(rec.score || 0);
                });
            } else {
                rawUnitMax = Number(unit.maxScore || unit.score || 10);
                const rec = scores.find(s => s.studentId === student.id && s.subjectId === subject.id && s.unitId === unit.id && !s.subUnitId);
                if (rec) rawUnitScore += Number(rec.score || 0);
            }

            // คำนวณทอนสัดส่วนให้อยู่ในโควตาคะแนนเก็บที่เหลือ
            let weightedUnitMax = totalRawUnitsMax > 0 ? (rawUnitMax / totalRawUnitsMax) * remainingQuota : 0;
            let weightedUnitScore = totalRawUnitsMax > 0 ? (rawUnitScore / totalRawUnitsMax) * remainingQuota : 0;

            totalWeightedUnitScore += weightedUnitScore;

            const statusTag = rawUnitScore > 0 
                ? `<span style="color: #16a34a; font-weight: bold;">✅ ส่งแล้ว</span>`
                : `<span style="color: #dc2626; font-weight: bold;">❌ ยังไม่ส่ง/ไม่มีคะแนน</span>`;

            rowsHtml += `
                <tr>
                    <td style="border: 1px solid #cbd5e1; padding: 6px 10px; font-weight: bold;">📖 ${unit.name}</td>
                    <td style="border: 1px solid #cbd5e1; padding: 6px; text-align: center; font-weight: bold; color: #2563eb;">${weightedUnitScore.toFixed(2)}</td>
                    <td style="border: 1px solid #cbd5e1; padding: 6px; text-align: center; color: #64748b;">${weightedUnitMax.toFixed(2)}</td>
                    <td style="border: 1px solid #cbd5e1; padding: 6px; text-align: center;">${statusTag}</td>
                </tr>`;
        });
    }

    // 5. ดึงคะแนนสอบกลางภาค / ปลายภาค
    const midRec = scores.find(s => s.studentId === student.id && s.subjectId === subject.id && (s.unitValue === 'midterm' || s.unitId === 'midterm'));
    const finalRec = scores.find(s => s.studentId === student.id && s.subjectId === subject.id && (s.unitValue === 'final' || s.unitId === 'final'));

    const midScore = midRec ? Number(midRec.score || 0) : 0;
    const finalScore = finalRec ? Number(finalRec.score || 0) : 0;

    // 6. คะแนนรวมทั้งหมด (คะแนนเก็บทอนสัดส่วน + กลางภาค + ปลายภาค)
    const finalTotalScore = totalWeightedUnitScore + midScore + finalScore;

    // คำนวณเกรด
    let grade = "0";
    if (finalTotalScore >= 80) grade = "4";
    else if (finalTotalScore >= 75) grade = "3.5";
    else if (finalTotalScore >= 70) grade = "3";
    else if (finalTotalScore >= 65) grade = "2.5";
    else if (finalTotalScore >= 60) grade = "2";
    else if (finalTotalScore >= 55) grade = "1.5";
    else if (finalTotalScore >= 50) grade = "1";

    cardContainer.innerHTML = `
        <div style="text-align: center; border-bottom: 2px solid #2563eb; padding-bottom: 8px; margin-bottom: 12px;">
            <h2 style="margin: 0; color: #1e3a8a; font-size: 20px;">รายงานสรุปผลการเรียนรายบุคคล</h2>
            <p style="margin: 3px 0 0 0; color: #475569; font-size: 13px;"><b>วิชา:</b> ${subject.code || ''} ${subject.name} (${subject.grade || ''})</p>
        </div>

        <div style="display: flex; justify-content: space-between; background: #f8fafc; padding: 10px 12px; border: 1px solid #e2e8f0; border-radius: 6px; margin-bottom: 12px; font-size: 13px;">
            <div>
                <span><b>ชื่อ-สกุล:</b> ${student.name}</span><br>
                <span><b>รหัสนักเรียน:</b> ${student.code || student.id || '-'}</span>
            </div>
            <div style="text-align: right;">
                <span><b>ห้องเรียน:</b> ${student.className || '-'}</span><br>
                <span><b>คะแนนรวม:</b> <b style="color: #2563eb; font-size: 15px;">${finalTotalScore.toFixed(2)}</b> / 100.00 | <b>เกรด:</b> <b style="color: ${grade === '0' ? '#ef4444' : '#16a34a'}; font-size: 16px;">${grade}</b></span>
            </div>
        </div>

        <h3 style="color: #0f172a; margin: 0 0 8px 0; font-size: 14px;">📊 คะแนนสรุปแยกตามบทเรียน (ทอนสัดส่วนรวม ${remainingQuota} คะแนน)</h3>
        <table style="width: 100%; border-collapse: collapse; font-size: 12px;">
            <thead>
                <tr style="background-color: #f1f5f9; color: #334155;">
                    <th style="border: 1px solid #cbd5e1; padding: 6px; text-align: left;">บทเรียน / รายการประเมิน</th>
                    <th style="border: 1px solid #cbd5e1; padding: 6px; text-align: center; width: 90px;">คะแนนที่ได้</th>
                    <th style="border: 1px solid #cbd5e1; padding: 6px; text-align: center; width: 90px;">คะแนนเต็ม</th>
                    <th style="border: 1px solid #cbd5e1; padding: 6px; text-align: center; width: 110px;">สถานะ</th>
                </tr>
            </thead>
            <tbody>
                ${rowsHtml || '<tr><td colspan="4" style="text-align:center; padding:10px; color:#94a3b8;">ไม่พบข้อมูลบทเรียน</td></tr>'}
                <tr style="background-color: #fefce8;">
                    <td style="border: 1px solid #cbd5e1; padding: 6px; font-weight: bold;">🎯 สอบกลางภาค</td>
                    <td style="border: 1px solid #cbd5e1; padding: 6px; text-align: center; font-weight: bold; color: #2563eb;">${midScore.toFixed(2)}</td>
                    <td style="border: 1px solid #cbd5e1; padding: 6px; text-align: center; color: #64748b;">${midMax.toFixed(2)}</td>
                    <td style="border: 1px solid #cbd5e1; padding: 6px; text-align: center;">-</td>
                </tr>
                <tr style="background-color: #fefce8;">
                    <td style="border: 1px solid #cbd5e1; padding: 6px; font-weight: bold;">🎯 สอบปลายภาค</td>
                    <td style="border: 1px solid #cbd5e1; padding: 6px; text-align: center; font-weight: bold; color: #2563eb;">${finalScore.toFixed(2)}</td>
                    <td style="border: 1px solid #cbd5e1; padding: 6px; text-align: center; color: #64748b;">${finalMax.toFixed(2)}</td>
                    <td style="border: 1px solid #cbd5e1; padding: 6px; text-align: center;">-</td>
                </tr>
            </tbody>
        </table>
    `;

    return cardContainer;
}

// ==============================
// แสดงตารางคะแนนรวมและรายงานงานค้าง
// ==============================

// ==============================
// แสดงตารางคะแนนรวมและรายงานงานค้าง (กรองตรงระดับชั้น)
// ==============================

// 1. สรุปคะแนนรวมและเกรด (แสดงเฉพาะนักเรียนที่มีระดับชั้นตรงกับวิชา)
function renderScores() {
    const table = document.getElementById("scoreTable");
    if (!table) return;

    table.innerHTML = "";

    students.forEach(student => {
        subjects.forEach(subject => {
            if (subject.grade && student.className && !student.className.startsWith(subject.grade)) {
                return;
            }

            const studentScores = scores.filter(s => s.studentId === student.id && s.subjectId === subject.id);
            if (studentScores.length === 0) return;

            const result = calculateSubjectTotal(student.id, subject.id);

            const row = document.createElement("tr");
            row.innerHTML = `
                <td>${student.code || "-"}</td>
                <td>${student.name} (${student.className})</td>
                <td><span style="background: #2563eb; color: white; padding: 1px 6px; border-radius: 4px; font-size: 11px;">${subject.grade || 'ทั่วไป'}</span> ${subject.code} - ${subject.name}</td>
                <td><b>${result.totalScore}</b> / 100</td>
                <td>${result.percent}%</td>
                <td><strong style="color: #16a34a; font-size: 16px;">${result.gradeNum}</strong> <span style="color: #64748b; font-size: 13px;">(${result.gradeLetter})</span></td>
            `;

            table.appendChild(row);
        });
    });
}

// 2. รายงานการขาดส่งงาน / งานค้างสะสม (แสดงเฉพาะนักเรียนที่มีระดับชั้นตรงกับวิชา)
function renderMissingTasks() {
    const missingTable = document.getElementById("missingTaskTable");
    const subjectId = Number(document.getElementById("scoreSubjectSelect")?.value);

    if (!missingTable) return;

    missingTable.innerHTML = "";

    // ถ้ายังไม่ได้เลือกวิชา ให้ซ่อนหรือแจ้งเตือนเลือกวิชาก่อน
    if (!subjectId) {
        missingTable.innerHTML = `<tr><td colspan="4" style="text-align: center; color: #6b7280; padding: 12px;">กรุณาเลือกวิชาด้านบนเพื่อดูรายงานงานค้าง</td></tr>`;
        return;
    }

    const subject = subjects.find(s => s.id === subjectId);
    if (!subject) return;

    students.forEach(student => {
        // 1. เช็กระดับชั้นของนักเรียนให้ตรงกับวิชาที่เลือก (เช่น ม.6 ตรงกับ ม.6/1)
        if (subject.grade && student.className && !student.className.startsWith(subject.grade)) {
            return; 
        }

        let missingItems = [];

        // 2. เช็กเฉพาะบทเรียนในวิชาที่กำลังเลือกอยู่
        if (subject.units) {
            subject.units.forEach(unit => {
                if (unit.subUnits && unit.subUnits.length > 0) {
                    unit.subUnits.forEach(sub => {
                        const isSubmitted = scores.some(
                            s => s.studentId === student.id && s.subjectId === subject.id && s.subUnitId === sub.id
                        );
                        if (!isSubmitted) {
                            missingItems.push(`[บท: ${unit.name}] - ${sub.name}`);
                        }
                    });
                } else {
                    const isSubmitted = scores.some(
                        s => s.studentId === student.id && s.subjectId === subject.id && s.unitId === unit.id
                    );
                    if (!isSubmitted) {
                        missingItems.push(`บท: ${unit.name}`);
                    }
                }
            });
        }

        if (missingItems.length > 0) {
            const row = document.createElement("tr");
            row.innerHTML = `
                <td style="padding: 8px;">${student.code || "-"}</td>
                <td style="padding: 8px;"><b>${student.name}</b> (${student.className})</td>
                <td style="padding: 8px;"><span style="background: #2563eb; color: white; padding: 1px 6px; border-radius: 4px; font-size: 11px;">${subject.grade || 'ทั่วไป'}</span> ${subject.code} - ${subject.name}</td>
                <td style="padding: 8px; color: #dc2626;">
                    ${missingItems.map(item => `<span style="display: inline-block; background: #fee2e2; padding: 2px 6px; border-radius: 4px; margin: 2px; font-size: 12px;">❌ ${item}</span>`).join(" ")}
                </td>
            `;
            missingTable.appendChild(row);
        }
    });

    if (missingTable.innerHTML === "") {
        missingTable.innerHTML = `<tr><td colspan="4" style="text-align: center; color: #10b981; padding: 12px;">🎉 ไม่มีงานค้างในวิชานี้ ทุกคนส่งงานครบถ้วน!</td></tr>`;
    }
}

// ==============================
// บันทึกและอัปเดตระบบ
// ==============================

function saveData() {
    localStorage.setItem("students", JSON.stringify(students));
    localStorage.setItem("subjects", JSON.stringify(subjects));
    localStorage.setItem("scores", JSON.stringify(scores));
}

function updateAll() {
    renderStudents();
    renderSubjects();
    renderSubjectSelectDropdown();
    renderScoreMatrix();
    renderScores();
    renderMissingTasks();

    if (document.getElementById("studentCount")) document.getElementById("studentCount").textContent = students.length;
    if (document.getElementById("subjectCount")) document.getElementById("subjectCount").textContent = subjects.length;
    if (document.getElementById("scoreCount")) document.getElementById("scoreCount").textContent = scores.length;
}
// ==============================
// ฟังก์ชันลบนักเรียนทีละเยอะๆ
// ==============================

// 1. ลบนักเรียนยกห้อง
function deleteStudentsByClass() {
    const targetClass = prompt("กรุณากรอกชื่อห้องที่ต้องการลบ (เช่น ม.3/1):");
    if (!targetClass) return;

    const trimmedClass = targetClass.trim();
    const studentsInClass = students.filter(s => s.className === trimmedClass);

    if (studentsInClass.length === 0) {
        alert(`ไม่พบข้อมูลนักเรียนในห้อง ${trimmedClass}`);
        return;
    }

    if (confirm(`คุณต้องการลบนักเรียนห้อง ${trimmedClass} ทั้งหมดจำนวน ${studentsInClass.length} คนใช่หรือไม่?\n(คะแนนของนักเรียนห้องนี้จะถูกลบไปด้วย)`)) {
        // ดึง ID ของนักเรียนในห้องนั้น
        const studentIdsToRemove = new Set(studentsInClass.map(s => s.id));

        // กรองเอาเฉพาะคนที่ไม่ได้อยู่ในห้องนั้นเก็บไว้
        students = students.filter(s => s.className !== trimmedClass);
        
        // ลบคะแนนที่เกี่ยวข้องกับนักเรียนกลุ่มนี้ออก
        scores = scores.filter(s => !studentIdsToRemove.has(s.studentId));

        saveData();
        updateAll();
        alert(`ลบนักเรียนห้อง ${trimmedClass} เรียบร้อยแล้ว`);
    }
}

// 2. ล้างข้อมูลนักเรียนทั้งหมดในระบบ
function clearAllStudents() {
    if (students.length === 0) {
        alert("ไม่มีข้อมูลนักเรียนในระบบ");
        return;
    }

    if (confirm("⚠️ ยืนยันการล้างข้อมูลนักเรียน 'ทั้งหมด' ในระบบ?\nข้อมูลนักเรียนและคะแนนทั้งหมดจะถูกลบออก และไม่สามารถกู้คืนได้!")) {
        students = [];
        scores = []; // ลบคะแนนทั้งหมดออกด้วย
        saveData();
        updateAll();
        alert("ล้างข้อมูลนักเรียนทั้งหมดเรียบร้อยแล้ว");
    }
}
// ==============================
// อัปเดต Dropdown เลือกรายบุคคล
// ==============================
function updateSingleStudentDropdown() {
    const select = document.getElementById("singleStudentSelect");
    const subjectId = Number(document.getElementById("scoreSubjectSelect")?.value);
    if (!select) return;

    select.innerHTML = `<option value="">-- เลือกรายบุคคล --</option>`;
    if (!subjectId) return;

    const subject = subjects.find(s => s.id === subjectId);
    
    // ดึงเฉพาะเด็กตรงระดับชั้น
    const filteredStudents = students.filter(student => {
        if (!subject || !subject.grade) return true;
        return student.className && student.className.startsWith(subject.grade);
    });

    filteredStudents.forEach(st => {
        const option = document.createElement("option");
        option.value = st.id;
        option.textContent = `${st.code || '-'} ${st.name} (${st.className})`;
        select.appendChild(option);
    });
}

// อัปเดต Dropdown เมื่อเปลี่ยนวิชา
const originalRenderScoreMatrix = renderScoreMatrix;
renderScoreMatrix = function() {
    originalRenderScoreMatrix();
    updateSingleStudentDropdown();
};


/// ==========================================
// ฟังก์ชันสร้างรายงาน PDF (สรุปคะแนนรวมแยกตามบทเรียน)
// ==========================================

// ==========================================
// ระบบส่งออกรายงาน PDF
// ==========================================

// 1. ส่งออกรายงานรายบุคคล
async function exportSingleStudentPDF() {
    const studentSelect = document.getElementById("singleStudentSelect") || document.getElementById("pdfStudentSelect");
    const subjectSelect = document.getElementById("scoreSubjectSelect");

    if (!subjectSelect || !subjectSelect.value) {
        alert("กรุณาเลือกรายวิชาก่อนครับ");
        return;
    }
    if (!studentSelect || !studentSelect.value) {
        alert("กรุณาเลือกนักเรียนที่ต้องการออกรายงานครับ");
        return;
    }

    const studentId = String(studentSelect.value).trim();
    const subjectId = String(subjectSelect.value).trim();

    const student = students.find(s => String(s.id).trim() === studentId || String(s.code).trim() === studentId);
    const subject = subjects.find(s => String(s.id).trim() === subjectId);

    if (!student || !subject) {
        alert("ไม่พบข้อมูลนักเรียนหรือรายวิชา");
        return;
    }

    const printArea = document.getElementById("pdfPrintArea");
    printArea.style.display = "block";
    printArea.innerHTML = ""; // ล้างหน้าเก่า

    const card = buildStudentReportCard(student, subject);
    printArea.appendChild(card);

    const opt = {
        margin:       [10, 10, 10, 10],
        filename:     `รายงานคะแนน_${student.name}_${subject.code || subject.name}.pdf`,
        image:        { type: 'jpeg', quality: 0.98 },
        html2canvas:  { scale: 2, useCORS: true, scrollY: 0 },
        jsPDF:        { unit: 'mm', format: 'a4', orientation: 'portrait' }
    };

    try {
        await html2pdf().set(opt).from(printArea).save();
    } catch (e) {
        console.error("PDF Export Error:", e);
    } finally {
        printArea.style.display = "none";
    }
}

// 2. ส่งออกรายงานทั้งชั้นเรียน
async function exportClassPDF() {
    const subjectSelect = document.getElementById("scoreSubjectSelect");

    if (!subjectSelect || !subjectSelect.value) {
        alert("กรุณาเลือกรายวิชาก่อนครับ");
        return;
    }

    const subjectId = String(subjectSelect.value).trim();
    const subject = subjects.find(s => String(s.id).trim() === subjectId);

    if (!subject) {
        alert("ไม่พบข้อมูลรายวิชา");
        return;
    }

    // กรองนักเรียนเฉพาะระดับชั้น
    const filteredStudents = students.filter(student => {
        if (!subject.grade) return true;
        return student.className && student.className.startsWith(subject.grade);
    });

    if (filteredStudents.length === 0) {
        alert("ไม่พบนักเรียนในระดับชั้นนี้");
        return;
    }

    const printArea = document.getElementById("pdfPrintArea");
    printArea.style.display = "block";
    printArea.innerHTML = "";

    for (let i = 0; i < filteredStudents.length; i++) {
        const student = filteredStudents[i];
        const card = buildStudentReportCard(student, subject);
        if (i < filteredStudents.length - 1) {
            card.style.pageBreakAfter = "always";
        }
        printArea.appendChild(card);
    }

    const opt = {
        margin:       [10, 10, 10, 10],
        filename:     `รายงานคะแนนชั้นเรียน_${subject.grade || ''}_${subject.code || subject.name}.pdf`,
        image:        { type: 'jpeg', quality: 0.98 },
        html2canvas:  { scale: 2, useCORS: true, scrollY: 0 },
        jsPDF:        { unit: 'mm', format: 'a4', orientation: 'portrait' }
    };

    try {
        await html2pdf().set(opt).from(printArea).save();
    } catch (e) {
        console.error("PDF Class Export Error:", e);
    } finally {
        printArea.style.display = "none";
    }
}
// 1. นำ URL จาก Google Apps Script มาวางตรงนี้
const GAs_API_URL = "https://script.google.com/macros/s/AKfycbxjuMkuZDfhOgThGScd6BaDHzXwmDkjJCaJ4f1LNONlyQ5fzfzZlneM6UTarylqd4s/exec";
// 2. ฟังก์ชันบันทึกคะแนนตรงลง Google Drive
async function saveMatrixScore(studentId, subjectId, unitKey, unitId, subUnitId, maxScore, inputElem) {
    const scoreVal = parseFloat(inputElem.value);

    if (scoreVal > maxScore) {
        alert(`คะแนนต้องไม่เกิน ${maxScore}`);
        inputElem.value = "";
        return;
    }

    inputElem.style.borderColor = "#f59e0b"; // สีเหลือง = กำลังส่งข้อมูล

    // อัปเดตในความจำเครื่องชั่วคราว
    let existingIndex = scores.findIndex(s => s.studentId === studentId && s.subjectId === subjectId && s.unitValue === unitKey);
    if (existingIndex > -1) {
        scores[existingIndex].score = scoreVal;
    } else {
        scores.push({
            studentId,
            subjectId,
            unitValue: unitKey,
            unitId,
            subUnitId,
            score: scoreVal
        });
    }

    // ยิงข้อมูลไป Google Sheets บน Drive
    try {
        await fetch(GAS_API_URL, {
            method: "POST",
            mode: "no-cors",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                studentId,
                subjectId,
                unitKey,
                unitId: unitId || "",
                subUnitId: subUnitId || "",
                score: isNaN(scoreVal) ? 0 : scoreVal
            })
        });

        inputElem.style.borderColor = "#10b981"; // สีเขียว = บันทึกลง Drive สำเร็จ!
        setTimeout(() => { inputElem.style.borderColor = "#d1d5db"; }, 1200);

    } catch (error) {
        console.error("การบันทึกล้มเหลว:", error);
        inputElem.style.borderColor = "#ef4444"; // สีแดง = เกิดข้อผิดพลาด
    }
}

// ==========================================
// 1. ฟังก์ชันโหลดข้อมูลนักเรียนและคะแนน
// ==========================================
// ==========================================
// ฟังก์ชันโหลดข้อมูลนักเรียนและคะแนนจาก Google Sheets (พร้อมดึงข้อมูลจริงขึ้นตาราง)
// ==========================================
async function loadScoresFromDrive() {
    // 1. อ่านค่าจาก localStorage ก่อน (ถ้ามี)
    const localStudents = localStorage.getItem('students');
    const localScores = localStorage.getItem('scores');

    if (localStudents) {
        try { students = JSON.parse(localStudents); } catch (e) {}
    }
    if (localScores) {
        try { scores = JSON.parse(localScores); } catch (e) {}
    }

    // เรนเดอร์รอบแรกจากความจำเบราว์เซอร์
    if (typeof renderScoreMatrix === 'function') renderScoreMatrix();
    if (typeof calculateAndRenderSummaryScores === 'function') calculateAndRenderSummaryScores();

    // 2. ดึงข้อมูลจริงทั้งหมดจาก Google Sheets ผ่าน GAS
    if (typeof GAS_API_URL !== 'undefined' && GAS_API_URL !== "") {
        try {
            // ใช้ mode: 'cors' หรือ query string เพื่อดึงข้อมูล JSON
            const response = await fetch(GAS_API_URL + "?action=getData");
            const data = await response.json();

            // ถ้ามีข้อมูลนักเรียนส่งกลับมาจาก Sheets ให้เอามาทับตัวแปรหลัก
            if (data && data.students && data.students.length > 0) {
                students = data.students;
                localStorage.setItem('students', JSON.stringify(students));
            }

            // ถ้ามีข้อมูลคะแนนส่งกลับมาจาก Sheets ให้เอามาทับตัวแปรหลัก
            if (data && data.scores && data.scores.length > 0) {
                scores = data.scores;
                localStorage.setItem('scores', JSON.stringify(scores));
            }

            // 3. บังคับวาดตารางและคะแนนใหม่อีกครั้งทันทีที่ได้ข้อมูลจาก Sheets
            if (typeof renderScoreMatrix === 'function') renderScoreMatrix();
            if (typeof calculateAndRenderSummaryScores === 'function') calculateAndRenderSummaryScores();

            console.log("โหลดข้อมูลจาก Google Sheets สำเร็จ:", { students, scores });
        } catch (e) {
            console.error("เกิดข้อผิดพลาดในการดึงข้อมูลจาก Google Sheets:", e);
        }
    }
}

// ==========================================
// 2. ฟังก์ชันเพิ่มนักเรียนใหม่
// ==========================================
async function addStudent(code, name, className) {
    if (!code || !name) {
        alert("กรุณากรอกรหัสและชื่อนักเรียนให้ครบถ้วน");
        return;
    }

    const newStudent = {
        id: String(code).trim(),
        code: String(code).trim(),
        name: String(name).trim(),
        className: className || "ม.6/1"
    };

    // เช็กว่ามีรหัสนักเรียนนี้อยู่แล้วหรือไม่
    const exists = students.some(s => String(s.id).trim() === String(newStudent.id).trim());
    if (exists) {
        alert("รหัสนักเรียนนี้มีอยู่ในระบบแล้ว");
        return;
    }

    // 1. เพิ่มเข้าอาร์เรย์ RAM
    students.push(newStudent);

    // 2. บันทึกลง localStorage ทันที
    localStorage.setItem('students', JSON.stringify(students));

    // 3. เรนเดอร์ตารางใหม่ทันที
    if (typeof renderScoreMatrix === 'function') renderScoreMatrix();
    if (typeof calculateAndRenderSummaryScores === 'function') calculateAndRenderSummaryScores();

    alert("เพิ่มนักเรียนเรียบร้อยแล้ว!");

    // 4. ส่งบันทึกลง Google Sheets
    if (typeof GAS_API_URL !== 'undefined' && GAS_API_URL !== "") {
        try {
            await fetch(GAS_API_URL, {
                method: "POST",
                body: JSON.stringify({
                    action: "addStudent",
                    student: newStudent
                })
            });
        } catch (e) {
            console.error("บันทึกนักเรียนลง Google Sheets ล้มเหลว:", e);
        }
    }
}
const GAS_API_URL = "https://script.google.com/macros/s/AKfycbxjuMkuZDfhOgThGScd6BaDHzXwmDkjJCaJ4f1LNONlyQ5fzfzZlneM6UTarylqd4s/exec";
// ==========================================
// ระบบบันทึก Session และควบคุมการเข้า/ออกจากระบบ
// ==========================================

// 1. ฟังก์ชันสลับสิทธิ์การเข้าใช้งานระหว่าง นักเรียน และ ครู หน้าแรก
function switchRole(role) {
    const studentForm = document.getElementById('studentLoginForm');
    const teacherForm = document.getElementById('teacherLoginForm');
    const btnStudent = document.getElementById('btnRoleStudent');
    const btnTeacher = document.getElementById('btnRoleTeacher');

    if (role === 'student') {
        if (studentForm) studentForm.style.display = 'block';
        if (teacherForm) teacherForm.style.display = 'none';
        if (btnStudent) btnStudent.classList.add('active');
        if (btnTeacher) btnTeacher.classList.remove('active');
    } else {
        if (studentForm) studentForm.style.display = 'none';
        if (teacherForm) teacherForm.style.display = 'block';
        if (btnTeacher) btnTeacher.classList.add('active');
        if (btnStudent) btnStudent.classList.remove('active');
    }
}

// 2. ฟังก์ชันล็อกอินครู (เรียกใช้ตอนกดปุ่มล็อกอินครู)
function loginTeacher() {
    // 💾 บันทึกสถานะ Session ไว้ใน LocalStorage
    localStorage.setItem('isLoggedIn', 'true');
    localStorage.setItem('userRole', 'teacher');

    const authContainer = document.getElementById('authContainer');
    const teacherView = document.getElementById('teacherSystemView');
    const studentView = document.getElementById('studentReportView');

    if (authContainer) authContainer.style.display = 'none';
    if (teacherView) teacherView.style.display = 'block';
    if (studentView) studentView.style.display = 'none';

    // โหลด/เรนเดอร์ข้อมูลตารางคะแนน
    if (typeof renderScoreMatrix === 'function') renderScoreMatrix();
    if (typeof calculateAndRenderSummaryScores === 'function') calculateAndRenderSummaryScores();
}

// 3. ฟังก์ชันล็อกอินนักเรียน (เรียกใช้ตอนกดปุ่มล็อกอินนักเรียน)
function loginStudent() {
    const studentCodeInput = document.getElementById('studentCodeInput');
    const studentCode = studentCodeInput ? studentCodeInput.value.trim() : '';

    if (!studentCode) {
        alert("กรุณากรอกรหัสนักเรียน");
        return;
    }

    // 💾 บันทึกสถานะ Session ไว้ใน LocalStorage
    localStorage.setItem('isLoggedIn', 'true');
    localStorage.setItem('userRole', 'student');
    localStorage.setItem('studentCode', studentCode);

    const authContainer = document.getElementById('authContainer');
    const teacherView = document.getElementById('teacherSystemView');
    const studentView = document.getElementById('studentReportView');

    if (authContainer) authContainer.style.display = 'none';
    if (teacherView) teacherView.style.display = 'none';
    if (studentView) studentView.style.display = 'block';
}

// 4. ปุ่มออกจากระบบ / กลับหน้าแรก (พร้อมลบการจำ Session)
function logout() {
    // 🧹 ลบสถานะการจำล็อกอินออกจากเครื่อง
    localStorage.removeItem('isLoggedIn');
    localStorage.removeItem('userRole');
    localStorage.removeItem('studentCode');

    const authContainer = document.getElementById('authContainer');
    const teacherView = document.getElementById('teacherSystemView');
    const studentView = document.getElementById('studentReportView');

    if (authContainer) authContainer.style.display = 'block';
    if (teacherView) teacherView.style.display = 'none';
    if (studentView) studentView.style.display = 'none';

    const studentCodeInput = document.getElementById('studentCodeInput');
    const teacherUser = document.getElementById('teacherUser');
    const teacherPass = document.getElementById('teacherPass');

    if (studentCodeInput) studentCodeInput.value = "";
    if (teacherUser) teacherUser.value = "";
    if (teacherPass) teacherPass.value = "";
}

// 5. ระบบตรวจสอบ Session ป้องกันการเด้งหลุดเมื่อ Refresh (DOMContentLoaded)
window.addEventListener("DOMContentLoaded", () => {
    // โหลดข้อมูลคะแนนจาก Drive ตามเดิม
    if (typeof loadScoresFromDrive === 'function') {
        loadScoresFromDrive();
    }

    // เช็กว่าเคยล็อกอินค้างไว้หรือไม่
    const isLoggedIn = localStorage.getItem('isLoggedIn');
    const userRole = localStorage.getItem('userRole');

    if (isLoggedIn === 'true') {
        const authContainer = document.getElementById('authContainer');
        const teacherView = document.getElementById('teacherSystemView');
        const studentView = document.getElementById('studentReportView');

        if (authContainer) authContainer.style.display = 'none';

        if (userRole === 'teacher') {
            if (teacherView) teacherView.style.display = 'block';
            if (studentView) studentView.style.display = 'none';
            if (typeof renderScoreMatrix === 'function') renderScoreMatrix();
        } else if (userRole === 'student') {
            if (studentView) studentView.style.display = 'block';
            if (teacherView) teacherView.style.display = 'none';
        }
    }
});

// 2. ฟังก์ชันตรวจสอบงานค้างของนักเรียน
async function checkStudentPendingTasks() {
    const inputElem = document.getElementById('studentCodeInput');
    if (!inputElem) return;
    
    const code = inputElem.value.trim();
    if (!code) {
        alert("กรุณากรอกเลขประจำตัวนักเรียน");
        return;
    }

    // ตรวจสอบว่ามีข้อมูลนักเรียนในระบบหรือไม่
    if (typeof students === 'undefined' || !Array.isArray(students) || students.length === 0) {
        alert("ยังไม่มีข้อมูลนักเรียนในระบบ");
        return;
    }

    // ค้นหานักเรียนจากเลขประจำตัว หรือ ID
    const student = students.find(s => 
        (s.code && String(s.code).trim() === code) || 
        (s.id && String(s.id).trim() === code)
    );

    if (!student) {
        alert(`ไม่พบข้อมูลนักเรียนที่มีเลขประจำตัว: ${code}`);
        return;
    }

    // โหลดคะแนนล่าสุดจาก Google Drive (ถ้ามีฟังก์ชัน)
    if (typeof loadScoresFromDrive === 'function') {
        try {
            await loadScoresFromDrive();
        } catch (e) {
            console.log("ไม่สามารถดึงข้อมูลคะแนนได้:", e);
        }
    }

    // แสดงผลรายงานสถานะการส่งงาน
    const reportContainer = document.getElementById('pendingTaskList');
    const titleElem = document.getElementById('studentReportTitle');
    
    if (titleElem) {
        titleElem.innerText = `รายงานสถานะการส่งงาน: ${student.name || ''} ${student.className ? '(' + student.className + ')' : ''}`;
    }

    if (reportContainer) {
        reportContainer.innerHTML = "";

        if (typeof subjects === 'undefined' || !Array.isArray(subjects) || subjects.length === 0) {
            reportContainer.innerHTML = "<p style='padding: 15px;'>ไม่พบรายการวิชาในระบบ</p>";
        } else {
            subjects.forEach(subject => {
                // กรองตามชั้นเรียนถ้ามีการกำหนดชั้น
                if (subject.grade && student.className && !student.className.startsWith(subject.grade)) return;

                let pendingHtml = `<div class="card" style="margin-bottom: 15px; text-align: left;"><h4>วิชา ${subject.code || ''} - ${subject.name || ''}</h4><ul style="line-height: 1.8;">`;
                
                if (subject.units && Array.isArray(subject.units)) {
                    subject.units.forEach(unit => {
                        const currentScores = typeof scores !== 'undefined' ? scores : [];
                        const isSubmitted = currentScores.some(s => 
                            String(s.studentId) === String(student.id) && 
                            String(s.subjectId) === String(subject.id) && 
                            (String(s.unitId) === String(unit.id) || String(s.unitValue) === String(unit.id)) && 
                            Number(s.score) > 0
                        );

                        if (!isSubmitted) {
                            pendingHtml += `<li style="color: #ef4444; font-weight: bold;">❌ ยังไม่ได้ส่ง: ${unit.name}</li>`;
                        } else {
                            pendingHtml += `<li style="color: #10b981;">✅ ส่งแล้ว: ${unit.name}</li>`;
                        }
                    });
                } else {
                    pendingHtml += `<li>ไม่มีรายการหน่วยการเรียนรู้</li>`;
                }

                pendingHtml += `</ul></div>`;
                reportContainer.innerHTML += pendingHtml;
            });
        }
    }

    // สลับหน้าจอแสดงผล
    const authBox = document.getElementById('authContainer');
    const reportBox = document.getElementById('studentReportView');
    
    if (authBox) authBox.style.display = 'none';
    if (reportBox) reportBox.style.display = 'block';
}

// 3. ฟังก์ชันออกจากระบบ / กลับหน้าหลัก
function logout() {
    const authBox = document.getElementById('authContainer');
    const reportBox = document.getElementById('studentReportView');
    const teacherBox = document.getElementById('teacherSystemView');

    if (authBox) authBox.style.display = 'block';
    if (reportBox) reportBox.style.display = 'none';
    if (teacherBox) teacherBox.style.display = 'none';

    const inputStudent = document.getElementById('studentCodeInput');
    const inputUser = document.getElementById('teacherUser');
    const inputPass = document.getElementById('teacherPass');

    if (inputStudent) inputStudent.value = "";
    if (inputUser) inputUser.value = "";
    if (inputPass) inputPass
    .value = "";
}
// ==========================================
// ระบบแก้ไขสัดส่วนคะแนนสอบกลางภาค / ปลายภาค
// ==========================================

// 1. ฟังก์ชันเปิดหน้าต่าง Modal สำหรับแก้ไขสัดส่วนคะแนน
function openEditSubjectModal(subjectId) {
    if (!subjectId) {
        alert("กรุณาเลือกรายวิชาที่ต้องการแก้ไขก่อนครับ");
        return;
    }

    const subject = subjects.find(s => String(s.id).trim() === String(subjectId).trim());
    if (!subject) {
        alert("ไม่พบข้อมูลรายวิชา");
        return;
    }

    // ดึงค่าเดิมที่มีอยู่ ถ้าไม่มีจะใช้ Default (กลางภาค 20, ปลายภาค 30)
    const midMax = subject.midtermMax || subject.midtermScore || 20;
    const finalMax = subject.finalMax || subject.finalScore || 30;
    const unitsQuota = 100 - midMax - finalMax;

    // เช็กว่ามี Modal เดิมอยู่หรือไม่ ถ้ามีให้ลบทิ้งก่อน
    closeEditSubjectModal();

    const modalHtml = `
    <div id="editSubjectModal" style="position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); display: flex; justify-content: center; align-items: center; z-index: 9999;">
        <div style="background: #fff; padding: 25px; border-radius: 12px; width: 90%; max-width: 420px; box-shadow: 0 10px 25px rgba(0,0,0,0.2); font-family: 'Sarabun', sans-serif;">
            <h3 style="margin-top: 0; color: #1e3a8a; border-bottom: 2px solid #2563eb; padding-bottom: 8px; font-size: 18px;">⚙️ แก้ไขสัดส่วนคะแนนสอบ</h3>
            
            <p style="font-weight: bold; color: #1e293b; margin-bottom: 15px; font-size: 14px;">วิชา: ${subject.code || ''} ${subject.name}</p>

            <div style="margin-bottom: 12px;">
                <label style="display: block; font-size: 13px; font-weight: bold; margin-bottom: 4px; color: #334155;">🎯 คะแนนเต็มสอบกลางภาค:</label>
                <input type="number" id="editMidtermScore" value="${midMax}" min="0" max="100" style="width: 100%; padding: 8px; border: 1px solid #cbd5e1; border-radius: 6px; box-sizing: border-box;" oninput="updateQuotaPreview()">
            </div>

            <div style="margin-bottom: 12px;">
                <label style="display: block; font-size: 13px; font-weight: bold; margin-bottom: 4px; color: #334155;">🎯 คะแนนเต็มสอบปลายภาค:</label>
                <input type="number" id="editFinalScore" value="${finalMax}" min="0" max="100" style="width: 100%; padding: 8px; border: 1px solid #cbd5e1; border-radius: 6px; box-sizing: border-box;" oninput="updateQuotaPreview()">
            </div>

            <div style="background: #f1f5f9; padding: 10px; border-radius: 6px; font-size: 13px; margin-bottom: 18px; color: #475569;">
                📖 คะแนนเก็บรวมทุกบทเรียน (ทอนให้อัตโนมัติ): <br>
                <strong id="quotaPreviewText" style="color: #2563eb; font-size: 15px;">${unitsQuota} คะแนน</strong>
            </div>

            <div style="display: flex; justify-content: flex-end; gap: 10px;">
                <button type="button" onclick="closeEditSubjectModal()" style="padding: 8px 16px; background: #94a3b8; color: #fff; border: none; border-radius: 6px; cursor: pointer;">ยกเลิก</button>
                <button type="button" onclick="saveSubjectExamSettings('${subject.id}')" style="padding: 8px 16px; background: #2563eb; color: #fff; border: none; border-radius: 6px; cursor: pointer; font-weight: bold;">💾 บันทึก</button>
            </div>
        </div>
    </div>`;

    document.body.insertAdjacentHTML('beforeend', modalHtml);
}

// 2. คำนวณตัวเลขพรีวิวขณะพิมพ์
function updateQuotaPreview() {
    const mid = Number(document.getElementById('editMidtermScore')?.value || 0);
    const final = Number(document.getElementById('editFinalScore')?.value || 0);
    const quota = 100 - mid - final;
    const quotaEl = document.getElementById('quotaPreviewText');
    if (quotaEl) {
        if (quota < 0) {
            quotaEl.style.color = '#ef4444';
            quotaEl.innerText = `เกินกำหนด (${quota} คะแนน)`;
        } else {
            quotaEl.style.color = '#2563eb';
            quotaEl.innerText = `${quota} คะแนน`;
        }
    }
}

// 3. ฟังก์ชันปิด Modal
function closeEditSubjectModal() {
    const modal = document.getElementById('editSubjectModal');
    if (modal) modal.remove();
}

// 4. บันทึกสัดส่วนคะแนนใหม่เข้าตัวแปรระบบ
async function saveSubjectExamSettings(subjectId) {
    const midVal = Number(document.getElementById('editMidtermScore').value || 0);
    const finalVal = Number(document.getElementById('editFinalScore').value || 0);

    if (midVal + finalVal > 100) {
        alert("ผลรวมคะแนนสอบกลางภาค + ปลายภาค ต้องไม่เกิน 100 คะแนนครับ");
        return;
    }

    const subject = subjects.find(s => String(s.id).trim() === String(subjectId).trim());
    if (subject) {
        subject.midtermMax = midVal;
        subject.midtermScore = midVal;
        subject.finalMax = finalVal;
        subject.finalScore = finalVal;

        // คำนวณและอัปเดตการแสดงผลตารางสดๆ ทันที
        if (typeof renderScoreMatrix === 'function') renderScoreMatrix();
        if (typeof calculateAndRenderSummaryScores === 'function') calculateAndRenderSummaryScores();

        closeEditSubjectModal();
        alert("อัปเดตสัดส่วนคะแนนสอบเรียบร้อยแล้วครับ!");

        // ส่งบันทึกไปยัง Google Apps Script (ถ้ามี)
        if (typeof GAS_API_URL !== 'undefined') {
            try {
                await fetch(GAS_API_URL, {
                    method: "POST",
                    body: JSON.stringify({
                        action: "updateSubjectExamMax",
                        subjectId: subjectId,
                        midtermMax: midVal,
                        finalMax: finalVal
                    })
                });
            } catch (e) {
                console.error("บันทึกลง Google Sheets ล้มเหลว:", e);
            }
        }
    }
}
// เริ่มต้นโปรแกรม
updateAll();