<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>سيستم إدارة السنتر التعليمي المطور</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <style>
        * { box-sizing: border-box; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
        body { background-color: #f4f6f9; margin: 0; padding: 15px; }
        .container { max-width: 1150px; margin: 0 auto; }
        h1 { text-align: center; color: #1a252f; margin-bottom: 20px; }
        .card { background: white; padding: 20px; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.07); margin-bottom: 20px; }
        .form-group { display: flex; gap: 10px; flex-wrap: wrap; }
        input, select, button { padding: 10px; border: 1px solid #ccc; border-radius: 6px; font-size: 15px; flex: 1; min-width: 140px; }
        button { cursor: pointer; font-weight: bold; border: none; transition: 0.2s; }
        .btn-add { background-color: #28a745; color: white; }
        .btn-excel { background-color: #1d6f42; color: white; width: 100%; margin-top: 10px; padding: 12px; font-size: 16px; }
        .btn-action { padding: 6px 10px; font-size: 12px; border-radius: 4px; margin: 2px; }
        .btn-present { background-color: #17a2b8; color: white; }
        .btn-pay { background-color: #ffc107; color: #333; }
        .btn-score { background-color: #6c757d; color: white; }
        .btn-whatsapp { background-color: #25D366; color: white; }
        .btn-delete { background-color: #dc3545; color: white; }
        table { width: 100%; border-collapse: collapse; margin-top: 15px; background: white; }
        th, td { border: 1px solid #dee2e6; padding: 10px; text-align: center; font-size: 13px; }
        th { background-color: #007bff; color: white; }
        tr:nth-child(even) { background-color: #f8f9fa; }
        .status-badge { padding: 4px 8px; border-radius: 4px; font-weight: bold; font-size: 11px; }
        .bg-success { background: #d4edda; color: #155724; }
        .bg-danger { background: #f8d7da; color: #721c24; }
        .bg-warning { background: #fff3cd; color: #856404; }
    </style>
</head>
<body>

<div class="container">
    <h1>لوحة إدارة السنتر التعليمي المطور</h1>

    <div class="card">
        <h3>تسجيل طالب جديد</h3>
        <div class="form-group">
            <input type="text" id="studentName" placeholder="اسم الطالب">
            <input type="text" id="studentPhone" placeholder="رقم الطالب">
            <input type="text" id="parentPhone" placeholder="رقم ولي الأمر">
            <select id="studentGrade">
                <option value="الأول الثانوي">الأول الثانوي</option>
                <option value="الثاني الثانوي">الثاني الثانوي</option>
                <option value="الثالث الثانوي">الثالث الثانوي</option>
            </select>
            <button class="btn-add" onclick="addStudent()">حفظ البيانات</button>
        </div>
    </div>

    <div class="card" style="overflow-x: auto;">
        <h3>بيانات وتقارير الطلاب</h3>
        <table>
            <thead>
                <tr>
                    <th>الكود</th>
                    <th>الاسم</th>
                    <th>الصف</th>
                    <th>الحضور</th>
                    <th>المصاريف</th>
                    <th>الدرجة</th>
                    <th>إجراءات وتحكم</th>
                    <th>إرسال تقرير</th>
                </tr>
            </thead>
            <tbody id="studentTable"></tbody>
        </table>
        <button class="btn-excel" onclick="exportToExcel()">تصدير كافة البيانات إلى ملف Excel 📊</button>
    </div>
</div>

<script>
    let students = JSON.parse(localStorage.getItem('center_students_v2')) || [];

    function saveAndRender() {
        localStorage.setItem('center_students_v2', JSON.stringify(students));
        renderTable();
    }

    function addStudent() {
        const name = document.getElementById('studentName').value.trim();
        const phone = document.getElementById('studentPhone').value.trim();
        const parentPhone = document.getElementById('parentPhone').value.trim();
        const grade = document.getElementById('studentGrade').value;

        if (!name || !phone) {
            alert('برجاء كتابة اسم الطالب ورقمه');
            return;
        }

        const newStudent = {
            id: students.length > 0 ? students[students.length - 1].id + 1 : 101,
            name, phone, parentPhone, grade,
            attended: false,
            paid: false,
            score: '-'
        };

        students.push(newStudent);
        saveAndRender();

        document.getElementById('studentName').value = '';
        document.getElementById('studentPhone').value = '';
        document.getElementById('parentPhone').value = '';
    }

    function toggleAttendance(index) {
        students[index].attended = !students[index].attended;
        saveAndRender();
    }

    function togglePayment(index) {
        students[index].paid = !students[index].paid;
        saveAndRender();
    }

    function updateScore(index) {
        const score = prompt('أدخل درجة الامتحان للطالب:', students[index].score);
        if (score !== null) {
            students[index].score = score;
            saveAndRender();
        }
    }

    function deleteStudent(index) {
        if (confirm('هل أنت تأكد من حذف هذا الطالب؟')) {
            students.splice(index, 1);
            saveAndRender();
        }
    }

    function sendWhatsApp(index) {
        const s = students[index];
        if (!s.parentPhone) {
            alert('لا يوجد رقم هاتف مضاف لولي الأمر!');
            return;
        }

        let phone = s.parentPhone.replace(/[^0-9]/g, '');
        if (phone.startsWith('0')) {
            phone = '2' + phone;
        }

        const attendText = s.attended ? 'حاضر ✅' : 'غائب ❌';
        const payText = s.paid ? 'تم التسديد 👍' : 'غير مسدد ⚠️';

        const message = `السلام عليكم ولي أمر الطالب/ة: *${s.name}*%0A` +
                        `تقرير متابعة السنتر التعليمي:%0A` +
                        `- حالة الحضور: ${attendText}%0A` +
                        `- درجة الامتحان: ${s.score}%0A` +
                        `- حالة المصاريف: ${payText}%0A` +
                        `شكراً لمتابعتكم معنا.`;

        window.open(`https://wa.me/${phone}?text=${message}`, '_blank');
    }

    function exportToExcel() {
        if (students.length === 0) {
            alert('لا يوجد بيانات طلاب لتصديرها!');
            return;
        }

        const excelData = students.map(s => ({
            'كود الطالب': s.id,
            'اسم الطالب': s.name,
            'الصف الدراسي': s.grade,
            'رقم الطالب': s.phone,
            'رقم ولي الأمر': s.parentPhone,
            'حالة الحضور': s.attended ? 'حاضر' : 'غائب',
            'حالة الدفع': s.paid ? 'تم الدفع' : 'مستحق',
            'الدرجة': s.score
        }));

        const worksheet = XLSX.utils.json_to_sheet(excelData);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "الطلاب");
        XLSX.writeFile(workbook, "تقرير_طلاب_السنتر.xlsx");
    }

    function renderTable() {
        const tbody = document.getElementById('studentTable');
        tbody.innerHTML = '';

        students.forEach((s, index) => {
            const attendStatus = s.attended 
                ? '<span class="status-badge bg-success">حاضر</span>' 
                : '<span class="status-badge bg-danger">غائب</span>';
                
            const payStatus = s.paid 
                ? '<span class="status-badge bg-success">تم الدفع</span>' 
                : '<span class="status-badge bg-warning">مستحق</span>';

            tbody.innerHTML += `
                <tr>
                    <td><b>${s.id}</b></td>
                    <td>${s.name}</td>
                    <td>${s.grade}</td>
                    <td>${attendStatus}</td>
                    <td>${payStatus}</td>
                    <td><b>${s.score}</b></td>
                    <td>
                        <button class="btn-action btn-present" onclick="toggleAttendance(${index})">الحضور</button>
                        <button class="btn-action btn-pay" onclick="togglePayment(${index})">الدفع</button>
                        <button class="btn-action btn-score" onclick="updateScore(${index})">الدرجة</button>
                        <button class="btn-action btn-delete" onclick="deleteStudent(${index})">حذف</button>
                    </td>
                    <td>
                        <button class="btn-action btn-whatsapp" onclick="sendWhatsApp(${index})">إرسال واتساب 📱</button>
                    </td>
                </tr>
            `;
        });
    }

    renderTable();
</script>

</body>
</html>
