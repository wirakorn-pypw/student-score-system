// ==============================
// ข้อมูลหลัก
// ==============================

let students = JSON.parse(localStorage.getItem("students")) || [];
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

function renderScoreMatrix() {
    const container = document.getElementById("matrixContainer");
    const subjectId = Number(document.getElementById("scoreSubjectSelect")?.value);

    if (!container) return;
    if (!subjectId) {
        container.innerHTML = `<p style="color: #6b7280;">กรุณาเลือกวิชาด้านบนเพื่อเริ่มกรอกคะแนน</p>`;
        return;
    }

    const subject = subjects.find(s => s.id === subjectId);
    if (!subject) {
        container.innerHTML = `<p style="color: #ef4444;">ไม่พบข้อมูลวิชา</p>`;
        return;
    }

    // กรองเฉพาะนักเรียนที่มีระดับชั้นตรงกับระดับชั้นของวิชา (เช่น "ม.6" จะตรงกับห้อง "ม.6/1", "ม.6/2")
    const filteredStudents = students.filter(student => {
        if (!subject.grade) return true; // ถ้าไม่ได้ระบุระดับชั้นวิชา ให้แสดงทุกคน
        return student.className && student.className.startsWith(subject.grade);
    });

    if (filteredStudents.length === 0) {
        container.innerHTML = `<p style="color: #ef4444;">ไม่พบรายชื่อนักเรียนในระดับชั้น ${subject.grade || 'นี้'} (กรุณาตรวจสอบชั้น/ห้องในหน้ารายชื่อนักเรียน)</p>`;
        return;
    }

    // รวบรวมคอลัมน์การสอบทั้งหมดของวิชานี้
    let columns = [];
    if (subject.midtermMax > 0) {
        columns.push({ key: "midterm", label: "🎯 กลางภาค", max: subject.midtermMax, unitId: null, subUnitId: null });
    }
    if (subject.finalMax > 0) {
        columns.push({ key: "final", label: "🏁 ปลายภาค", max: subject.finalMax, unitId: null, subUnitId: null });
    }

    if (subject.units) {
        subject.units.forEach(unit => {
            if (unit.subUnits && unit.subUnits.length > 0) {
                unit.subUnits.forEach(sub => {
                    columns.push({
                        key: `${unit.id}_${sub.id}`,
                        label: `📖 ${unit.name} - ${sub.name}`,
                        max: sub.maxScore,
                        unitId: unit.id,
                        subUnitId: sub.id
                    });
                });
            } else {
                columns.push({
                    key: `${unit.id}_none`,
                    label: `📖 บท: ${unit.name}`,
                    max: unit.maxScore,
                    unitId: unit.id,
                    subUnitId: null
                });
            }
        });
    }

    // สร้างตาราง Matrix
    let html = `
        <div style="margin-bottom: 10px; font-weight: bold; color: #1d4ed8;">
            📌 กำลังแสดงรายชื่อนักเรียนระดับชั้น: <span style="background: #2563eb; color: white; padding: 2px 8px; border-radius: 4px;">${subject.grade || 'แสดงทั้งหมด'}</span> (จำนวน ${filteredStudents.length} คน)
        </div>
        <table border="1" style="width: 100%; border-collapse: collapse; font-size: 14px; text-align: center;">
            <thead>
                <tr style="background-color: #f3f4f6;">
                    <th style="padding: 10px; min-width: 60px;">รหัส</th>
                    <th style="padding: 10px; min-width: 150px; text-align: left;">ชื่อ-สกุล</th>
                    <th style="padding: 10px; min-width: 80px;">ห้อง</th>
                    ${columns.map(col => `
                        <th style="padding: 10px; min-width: 110px;">
                            ${col.label}<br>
                            <span style="font-size: 11px; color: #2563eb;">(เต็ม ${col.max})</span>
                        </th>
                    `).join('')}
                </tr>
            </thead>
            <tbody>
    `;

    filteredStudents.forEach(student => {
        html += `
            <tr>
                <td style="padding: 8px;">${student.code || "-"}</td>
                <td style="padding: 8px; text-align: left;"><b>${student.name}</b></td>
                <td style="padding: 8px;">${student.className}</td>
        `;

        columns.forEach(col => {
            const currentScoreObj = scores.find(
                s => s.studentId === student.id && s.subjectId === subject.id && s.unitValue === col.key
            );
            const val = currentScoreObj !== undefined ? currentScoreObj.score : "";

            html += `
                <td style="padding: 5px;">
                    <input 
                        type="number" 
                        step="0.1"
                        min="0"
                        max="${col.max}"
                        value="${val}" 
                        placeholder="-"
                        style="width: 70px; text-align: center; padding: 6px; border: 1px solid #d1d5db; border-radius: 4px;"
                        onchange="saveMatrixScore(${student.id}, ${subject.id}, '${col.key}', ${col.unitId}, ${col.subUnitId}, ${col.max}, this)"
                    />
                </td>
            `;
        });

        html += `</tr>`;
    });

    html += `
            </tbody>
        </table>
        <p style="font-size: 12px; color: #10b981; margin-top: 10px;">✨ ระบบจะบันทึกคะแนนลงฐานข้อมูลให้อัตโนมัติทันทีที่พิมพ์หรือกด Tab เปลี่ยนช่อง</p>
    `;

    container.innerHTML = html;
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

function calculateSubjectTotal(studentId, subjectId) {
    const subject = subjects.find(s => s.id === subjectId);
    if (!subject) return { totalScore: 0, percent: 0, grade: "0", gradeLetter: "F" };

    const studentScores = scores.filter(s => s.studentId === studentId && s.subjectId === subjectId);
    let totalScore = 0;

    const midterm = studentScores.find(s => s.unitValue === "midterm");
    if (midterm) totalScore += midterm.score;

    const final = studentScores.find(s => s.unitValue === "final");
    if (final) totalScore += final.score;

    if (subject.units) {
        subject.units.forEach(unit => {
            if (unit.subUnits && unit.subUnits.length > 0) {
                unit.subUnits.forEach(sub => {
                    const subScoreRecord = studentScores.find(s => s.subUnitId === sub.id);
                    if (subScoreRecord) {
                        totalScore += subScoreRecord.score;
                    }
                });
            } else {
                const unitScoreRecord = studentScores.find(s => s.unitId === unit.id && !s.subUnitId);
                if (unitScoreRecord) {
                    totalScore += unitScoreRecord.score;
                }
            }
        });
    }

    const percent = Math.min(100, totalScore);
    
    // คำนวณเกรดตัวเลขและเกรดตัวอักษร
    let gradeNum = "0";
    let gradeLetter = "F";

    if (percent >= 80) { gradeNum = "4";   gradeLetter = "A";  }
    else if (percent >= 75) { gradeNum = "3.5"; gradeLetter = "B+"; }
    else if (percent >= 70) { gradeNum = "3";   gradeLetter = "B";  }
    else if (percent >= 65) { gradeNum = "2.5"; gradeLetter = "C+"; }
    else if (percent >= 60) { gradeNum = "2";   gradeLetter = "C";  }
    else if (percent >= 55) { gradeNum = "1.5"; gradeLetter = "D+"; }
    else if (percent >= 50) { gradeNum = "1";   gradeLetter = "D";  }
    else { gradeNum = "0"; gradeLetter = "F"; }

    return {
        totalScore: totalScore.toFixed(2),
        percent: percent.toFixed(1),
        gradeNum: gradeNum,         // เกรดตัวเลข (0, 1, 1.5, 2, 2.5, 3, 3.5, 4)
        gradeLetter: gradeLetter,   // เกรดตัวอักษร (F, D, D+, C, C+, B, B+, A)
        gradeFull: `${gradeNum} (${gradeLetter})` // แสดงคู่กัน เช่น 4 (A)
    };
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


// ==============================
// ฟังก์ชันสร้างรายงาน PDF พร้อมกราฟเรดาร์
// ==============================

// 1. ส่งออกรายงานรายคน
async function exportSingleStudentPDF() {
    const studentId = Number(document.getElementById("singleStudentSelect")?.value);
    const subjectId = Number(document.getElementById("scoreSubjectSelect")?.value);

    if (!subjectId) {
        alert("กรุณาเลือกรายวิชาก่อนครับ");
        return;
    }
    if (!studentId) {
        alert("กรุณาเลือกนักเรียนที่ต้องการออกรายงานครับ");
        return;
    }

    const student = students.find(s => s.id === studentId);
    const subject = subjects.find(s => s.id === subjectId);
    const printArea = document.getElementById("pdfPrintArea");

    printArea.style.display = "block";
    printArea.innerHTML = ""; // ล้างค่าเก่า

    const card = await buildStudentReportCard(student, subject, 0);
    printArea.appendChild(card);

    const opt = {
        margin:       10,
        filename:     `รายงานคะแนน_${student.name}_${subject.code}.pdf`,
        image:        { type: 'jpeg', quality: 0.98 },
        html2canvas:  { scale: 2, useCORS: true },
        jsPDF:        { unit: 'mm', format: 'a4', orientation: 'portrait' }
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
        const card = await buildStudentReportCard(student, subject, i);
        printArea.appendChild(card);
    }

    const opt = {
        margin:       8,
        filename:     `รายงานคะแนนชั้นเรียน_${subject.grade}_${subject.code}.pdf`,
        image:        { type: 'jpeg', quality: 0.98 },
        html2canvas:  { scale: 2, useCORS: true },
        jsPDF:        { unit: 'mm', format: 'a4', orientation: 'portrait' }
    };

    await html2pdf().set(opt).from(printArea).save();
    printArea.style.display = "none";
}


// ==============================
// ฟังก์ชันสร้างการ์ดรายงาน + กราฟ Radar
// ==============================

function buildStudentReportCard(student, subject, index) {
    return new Promise((resolve) => {
        const result = calculateSubjectTotal(student.id, subject.id);
        const cardContainer = document.createElement("div");
        cardContainer.style.cssText = "page-break-after: always; padding: 20px; font-family: sans-serif;";

        // เตรียมข้อมูลกราฟ Radar (ตามบทเรียน)
        const labels = [];
        const studentScores = [];
        const maxScores = [];

        if (subject.units && subject.units.length > 0) {
            subject.units.forEach(unit => {
                labels.push(unit.name);
                let scoreEarned = 0;

                if (unit.subUnits && unit.subUnits.length > 0) {
                    unit.subUnits.forEach(sub => {
                        const rec = scores.find(s => s.studentId === student.id && s.subjectId === subject.id && s.subUnitId === sub.id);
                        if (rec) scoreEarned += rec.score;
                    });
                } else {
                    const rec = scores.find(s => s.studentId === student.id && s.subjectId === subject.id && s.unitId === unit.id && !s.subUnitId);
                    if (rec) scoreEarned += rec.score;
                }

                studentScores.push(scoreEarned);
                maxScores.push(unit.maxScore || 0);
            });
        }

        const canvasId = `radarCanvas_${student.id}_${index}`;

        cardContainer.innerHTML = `
            <div style="text-align: center; border-bottom: 2px solid #2563eb; padding-bottom: 10px; margin-bottom: 15px;">
                <h2 style="margin: 0; color: #1e3a8a;">รายงานสรุปผลการเรียนและพัฒนาการ</h2>
                <p style="margin: 5px 0 0 0; color: #475569;">วิชา ${subject.code} - ${subject.name} (${subject.grade || ''})</p>
            </div>

            <div style="display: flex; justify-content: space-between; background: #f8fafc; padding: 12px; border-radius: 8px; margin-bottom: 15px; font-size: 14px;">
                <div>
                    <span><b>ชื่อ-สกุล:</b> ${student.name}</span><br>
                    <span><b>รหัสนักเรียน:</b> ${student.code || '-'}</span>
                </div>
                <div style="text-align: right;">
                    <span><b>ห้องเรียน:</b> ${student.className}</span><br>
<span><b>คะแนนรวม:</b> <b style="color: #2563eb;">${result.totalScore}</b> / 100 | <b>เกรด:</b> <b style="color: #16a34a; font-size: 16px;">${result.gradeNum}</b> (${result.gradeLetter})</span>
                </div>
            </div>

            <div style="display: flex; align-items: center; justify-content: center; gap: 20px;">
                <div style="width: 320px; height: 320px;">
                    <canvas id="${canvasId}"></canvas>
                </div>
            </div>
        `;

        document.body.appendChild(cardContainer); // แปะชั่วคราวเพื่อเรนเดอร์ Canvas

        setTimeout(() => {
            const ctx = document.getElementById(canvasId).getContext('2d');
            new Chart(ctx, {
                type: 'radar',
                data: {
                    labels: labels.length > 0 ? labels : ['กลางภาค', 'ปลายภาค'],
                    datasets: [{
                        label: 'คะแนนที่ได้',
                        data: studentScores.length > 0 ? studentScores : [
                            (scores.find(s=>s.studentId===student.id && s.subjectId===subject.id && s.unitValue==='midterm')?.score || 0),
                            (scores.find(s=>s.studentId===student.id && s.subjectId===subject.id && s.unitValue==='final')?.score || 0)
                        ],
                        backgroundColor: 'rgba(37, 99, 235, 0.2)',
                        borderColor: '#2563eb',
                        pointBackgroundColor: '#2563eb'
                    }, {
                        label: 'คะแนนเต็มบท',
                        data: maxScores.length > 0 ? maxScores : [subject.midtermMax || 0, subject.finalMax || 0],
                        backgroundColor: 'rgba(203, 213, 225, 0.2)',
                        borderColor: '#94a3b8',
                        borderDash: [5, 5]
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: {
                        r: { beginAtZero: true }
                    },
                    plugins: {
                        legend: { position: 'bottom' }
                    }
                }
            });

            setTimeout(() => {
                resolve(cardContainer);
            }, 300);
        }, 100);
    });
}
// 1. นำ URL จาก Google Apps Script มาวางตรงนี้
const GAS_API_URL = "https://script.googleusercontent.com/macros/echo?user_content_key=AUkAhnTiq7_72AjKFDfPLVBqI9T0Oqa-9wdNSAOXviWQ8Up2a9x1WIEIqNstkkWdKVvztzNpz9CxvQolGBtbl9iWTYCv5YFmEUzMsFaLvX0pYW9G0EHxyci8xY7jk1gHPPaZWllOS6RscIWLqCU1IJFoGhdGlT3xxKRTybRBSGbYZIxR6E7d791xe1-Hv5WMzP1HCcJnb85xiJ6k91IIxiIZfH3Vuuk78UdOl9FTtKn4avDogYY-STgifIx_lX1vOOxp5Qvi-1gFCu-TZPcel3LWHovzD2L9Bw&lib=M0H6212ObtVxveWcmtSYgQLlCWkN5VxJb"; 

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

// 3. ดึงคะแนนเก่าจาก Google Drive มาแสดงเมื่อเปิดเว็บ
async function loadScoresFromDrive() {
    try {
        const res = await fetch(GAS_API_URL);
        const data = await res.json();
        if (Array.isArray(data)) {
            scores = data.map(item => ({
                studentId: Number(item.studentId),
                subjectId: Number(item.subjectId),
                unitValue: item.unitKey,
                unitId: item.unitId ? Number(item.unitId) : null,
                subUnitId: item.subUnitId ? Number(item.subUnitId) : null,
                score: Number(item.score)
            }));
            if (typeof renderScoreMatrix === "function") renderScoreMatrix();
        }
    } catch (e) {
        console.log("ยังไม่มีข้อมูลเก่าใน Drive หรือโหลดไม่สำเร็จ:", e);
    }
}

window.addEventListener("DOMContentLoaded", loadScoresFromDrive);
// เริ่มต้นโปรแกรม
updateAll();