const API_URL = "https://cis-attendance.onrender.com";


// =====================================================
// CREATE SESSION
// =====================================================

async function createSession() {

    const eventName =
        document.getElementById("eventName").value.trim();

    const message =
        document.getElementById("sessionMessage");


    if (!eventName) {

        message.innerHTML =
            `<div class="error">
                Enter event name.
            </div>`;

        return;
    }


    try {

        const response = await fetch(
            API_URL + "/sessions",
            {
                method: "POST",

                headers: {
                    "Content-Type": "application/json"
                },

                body: JSON.stringify({
                    event: eventName
                })
            }
        );


        const data = await response.json();


        if (!response.ok) {

            message.innerHTML =
                `<div class="error">
                    ${data.detail}
                </div>`;

            return;
        }


        localStorage.setItem(
            "sessionId",
            data.session_id
        );


        localStorage.setItem(
            "sessionEvent",
            data.event
        );


        document.getElementById(
            "currentSession"
        ).textContent = data.session_id;


        document.getElementById(
            "currentEvent"
        ).textContent = data.event;


        message.innerHTML =
            `<div class="success">
                ✅ Attendance session started.
                <br>
                Session ID: ${data.session_id}
            </div>`;


        loadAttendance();

    } catch (error) {

        console.error(error);

        message.innerHTML =
            `<div class="error">
                ❌ Cannot connect to backend.
            </div>`;
    }
}


// =====================================================
// LOAD CURRENT SESSION
// =====================================================

function loadCurrentSession() {

    const sessionId =
        localStorage.getItem("sessionId");

    const event =
        localStorage.getItem("sessionEvent");


    if (sessionId) {

        const element =
            document.getElementById("currentSession");

        if (element) {

            element.textContent = sessionId;
        }
    }


    if (event) {

        const element =
            document.getElementById("currentEvent");

        if (element) {

            element.textContent = event;
        }
    }
}


// =====================================================
// LOAD ATTENDANCE
// =====================================================

async function loadAttendance() {

    const sessionId =
        localStorage.getItem("sessionId");


    if (!sessionId) {

        return;
    }


    try {

        const response = await fetch(
            API_URL +
            "/attendance/" +
            encodeURIComponent(sessionId)
        );


        const data = await response.json();


        if (!response.ok) {

            return;
        }


        document.getElementById(
            "totalStudents"
        ).textContent =
            data.total_students;


        document.getElementById(
            "presentStudents"
        ).textContent =
            data.present;


        document.getElementById(
            "absentStudents"
        ).textContent =
            data.absent;


        const list =
            document.getElementById(
                "attendanceList"
            );


        if (data.attendance.length === 0) {

            list.innerHTML =
                "No attendance marked yet.";

            return;
        }


        list.innerHTML = "";


        data.attendance.forEach(
            function(student) {

                const row =
                    document.createElement("div");

                row.className =
                    "attendance-row";


                row.innerHTML = `

                    <span>
                        <strong>
                            ${student.name}
                        </strong>
                        <br>
                        ${student.student_id}
                    </span>

                    <span>
                        ${student.time}
                        <br>
                        <strong>
                            ${student.status}
                        </strong>
                    </span>

                `;


                list.appendChild(row);

            }
        );


    } catch (error) {

        console.error(error);
    }
}


// =====================================================
// DOWNLOAD EXCEL
// =====================================================

function downloadExcel() {

    window.open(
        API_URL + "/download",
        "_blank"
    );
}


// =====================================================
// PAGE LOAD
// =====================================================

document.addEventListener(
    "DOMContentLoaded",
    function() {

        loadCurrentSession();

        loadAttendance();

    }
);