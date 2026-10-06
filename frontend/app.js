// =====================================================
// IEEE CIS QR ATTENDANCE SYSTEM
// APP.JS
// =====================================================


// =====================================================
// API
// =====================================================

const API_URL =
    "https://ieee-cis-qr-attendance.onrender.com";


// =====================================================
// LOGIN
// =====================================================

async function login() {

    const userIdElement =
        document.getElementById("userId");

    const passwordElement =
        document.getElementById("password");

    const messageElement =
        document.getElementById("message");


    if (!userIdElement || !passwordElement) {

        console.error(
            "Login input elements not found."
        );

        return;
    }


    const userId =
        userIdElement.value.trim();

    const password =
        passwordElement.value;


    if (!userId || !password) {

        if (messageElement) {

            messageElement.innerHTML = `
                <div class="error">
                    Please enter User ID and Password.
                </div>
            `;
        }

        return;
    }


    try {

        if (messageElement) {

            messageElement.innerHTML = `
                <div>
                    Logging in...
                </div>
            `;
        }


        const response =
            await fetch(
                `${API_URL}/login`,
                {
                    method: "POST",

                    headers: {
                        "Content-Type":
                            "application/json"
                    },

                    body:
                        JSON.stringify({
                            user_id: userId,
                            password: password
                        })
                }
            );


        let data = {};

        try {

            data =
                await response.json();

        } catch (error) {

            data = {};
        }


        if (!response.ok) {

            throw new Error(
                data.detail ||
                "Invalid User ID or Password."
            );
        }


        // =================================================
        // SAVE LOGIN INFORMATION
        // =================================================

        sessionStorage.setItem(
            "authToken",
            data.token
        );

        sessionStorage.setItem(
            "userId",
            data.user_id
        );

        sessionStorage.setItem(
            "userName",
            data.name
        );

        sessionStorage.setItem(
            "userRole",
            data.role
        );


        // =================================================
        // REDIRECT BASED ON ROLE
        // =================================================

        if (data.role === "HEAD") {

            window.location.href =
                "dashboard.html";

            return;
        }


        if (
            data.role ===
            "SESSION_CREATOR"
        ) {

            window.location.href =
                "dashboard.html";

            return;
        }


        if (
            data.role ===
            "STUDENT"
        ) {

            window.location.href =
                "student-dashboard.html";

            return;
        }


        // Unknown role

        sessionStorage.clear();

        throw new Error(
            "Unknown user role."
        );


    } catch (error) {

        console.error(
            "Login error:",
            error
        );


        if (messageElement) {

            messageElement.innerHTML = `
                <div class="error">
                    ❌ ${escapeHtml(
                        error.message
                    )}
                </div>
            `;
        }
    }
}


// =====================================================
// LOGOUT
// =====================================================

async function logout() {

    const token =
        sessionStorage.getItem(
            "authToken"
        );


    try {

        if (token) {

            await fetch(
                `${API_URL}/logout`,
                {
                    method: "POST",

                    headers: {
                        "Authorization":
                            `Bearer ${token}`
                    }
                }
            );
        }

    } catch (error) {

        console.error(
            "Logout error:",
            error
        );
    }


    sessionStorage.clear();


    window.location.href =
        "index.html";
}


// =====================================================
// CHECK LOGIN
// =====================================================

function isLoggedIn() {

    const token =
        sessionStorage.getItem(
            "authToken"
        );

    return !!token;
}


// =====================================================
// GET CURRENT ROLE
// =====================================================

function getUserRole() {

    return sessionStorage.getItem(
        "userRole"
    );
}


// =====================================================
// GET CURRENT USER ID
// =====================================================

function getUserId() {

    return sessionStorage.getItem(
        "userId"
    );
}


// =====================================================
// GET CURRENT USER NAME
// =====================================================

function getUserName() {

    return sessionStorage.getItem(
        "userName"
    );
}


// =====================================================
// PROTECT PAGE
// =====================================================

function requireLogin() {

    if (!isLoggedIn()) {

        window.location.href =
            "index.html";

        return false;
    }

    return true;
}


// =====================================================
// PROTECT HEAD PAGE
// =====================================================

function requireHead() {

    const token =
        sessionStorage.getItem(
            "authToken"
        );

    const role =
        sessionStorage.getItem(
            "userRole"
        );


    if (
        !token ||
        role !== "HEAD"
    ) {

        window.location.href =
            "index.html";

        return false;
    }


    return true;
}


// =====================================================
// PROTECT SESSION CREATOR PAGE
// =====================================================

function requireSessionCreator() {

    const token =
        sessionStorage.getItem(
            "authToken"
        );

    const role =
        sessionStorage.getItem(
            "userRole"
        );


    if (
        !token ||
        (
            role !== "HEAD" &&
            role !== "SESSION_CREATOR"
        )
    ) {

        window.location.href =
            "index.html";

        return false;
    }


    return true;
}


// =====================================================
// PROTECT STUDENT PAGE
// =====================================================

function requireStudent() {

    const token =
        sessionStorage.getItem(
            "authToken"
        );

    const role =
        sessionStorage.getItem(
            "userRole"
        );


    if (
        !token ||
        role !== "STUDENT"
    ) {

        window.location.href =
            "index.html";

        return false;
    }


    return true;
}


// =====================================================
// HTML SAFETY
// =====================================================

function escapeHtml(value) {

    if (
        value === null ||
        value === undefined
    ) {

        return "";
    }


    return String(value)
        .replace(
            /&/g,
            "&amp;"
        )
        .replace(
            /</g,
            "&lt;"
        )
        .replace(
            />/g,
            "&gt;"
        )
        .replace(
            /"/g,
            "&quot;"
        )
        .replace(
            /'/g,
            "&#039;"
        );
}


// =====================================================
// AUTO LOGIN FORM SUPPORT
// =====================================================

document.addEventListener(
    "DOMContentLoaded",
    function () {

        const loginForm =
            document.getElementById(
                "loginForm"
            );


        if (loginForm) {

            loginForm.addEventListener(
                "submit",
                function (event) {

                    event.preventDefault();

                    login();
                }
            );
        }

    }
);