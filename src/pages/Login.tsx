import React from 'react';
import '../assets/styles/Login.css';
import { msalInstance } from '../auth/msalInstance';

const Login = () => {
    const handleLogin = () => {
        msalInstance.loginPopup({
            scopes: ["User.Read"],
        }).then((response) => {
            console.log("Login successful:", response);
        }).catch((error) => {
            console.error("Login failed:", error);
        });
    };

    return (
        <div className="Login">
            <h1>Power BI Embed Demo</h1>
            <button onClick={handleLogin}>Login</button>
        </div>
    );
};

export default Login;