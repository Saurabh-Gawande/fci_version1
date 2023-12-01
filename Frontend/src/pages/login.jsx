import React, { useEffect, useState } from "react";
import "./Login.css";
import config from "../config";

function Login() {
  const [username, setUsername] = useState();
  const [password, setPassword] = useState();
  const [currentImage, setCurrentImage] = useState(0);
  const [state, setState] = useState();

  const images = ["static/img/slider1.jpg", "static/img/slider6.jpg"];
  const ProjectIp = config.serverUrl;
  const totalImages = images.length;
  const autoSlideInterval = 3000; // Interval in milliseconds (5 seconds in this example)

  useEffect(() => {
    const timer = setInterval(() => {
      setCurrentImage((prevImage) => (prevImage + 1) % totalImages);
    }, autoSlideInterval);

    return () => {
      clearInterval(timer);
    };
  }, [totalImages]);

  const handleLogin = async () => {
    try {
      const form = new FormData();
      form.append("username", username);
      form.append("password", password);

      const response = await fetch(ProjectIp + "/login", {
        method: "POST",
        credentials: "include",
        headers: {
          "Content-Type": "application/json",
        },
        // body: form,
        body: JSON.stringify({ username, password }),
      });
      const data = await response.json();

      if (response.ok) {
        setState(data.state);
        sessionStorage.setItem("state", data.state);
        window.alert("Login Successful! Click Ok to Continue");
        if (data.state === "Fcihq") {
          window.location.assign("/Monthly_Solution");
        } else {
          window.location.assign("/Daily_Planner");
        }
      } else {
        window.alert("Invallid Credentials");
        window.location.reload();
      }
    } catch (error) {
      console.error("Error during login:", error);
      window.alert("An error occurred during login. Please try again later.");
    }
  };
 
  return (
    <div
      className="login-container login_container"
      style={{
        backgroundImage: `url(${images[currentImage]})`,
        height: "100vh",
      }}
    >
      <div className="content">
        <div className="login-box animated fadeInDown">
          <div className="img-responsive">
            <center>
              <img src="static/img/FCI.jpg" width="300" alt="" />
            </center>
          </div>
          <br />
          <div className="login-body">
            <div className="login-title" style={{ color: "white" }}>
              <strong>Welcome</strong>, Please login
            </div>
            <form className="form-horizontal">
              <div className="form-group">
                <div className="col-md-12">
                  <input
                    type="text"
                    id="username"
                    name="username"
                    className="form-control"
                    placeholder="Username"
                    style={{ backgroundColor: "white", color: "black" }}
                    onChange={(e) => setUsername(e.target.value)}
                  />
                </div>
              </div>
              <div className="form-group">
                <div className="col-md-12">
                  <input
                    style={{ backgroundColor: "white" }}
                    type="password"
                    id="password"
                    name="password"
                    className="form-control"
                    placeholder="Password"
                    onChange={(e) => setPassword(e.target.value)}
                  />
                </div>
              </div>
              <div className="form-group">
                <center>
                  <input
                    type="button"
                    className="btn btn-primary"
                    value="login"
                    onClick={handleLogin}
                    style={{ borderRadius: "5px" }}
                  ></input>
                </center>
              </div>
            </form>
          </div>
          <div className="login-footer">
            <div className="pull-left" style={{ color: "white" }}>
              &copy; DEVELOPED BY IIT-DELHI v2.0
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

export default Login;
