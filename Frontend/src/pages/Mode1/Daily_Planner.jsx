import React, { useRef, useState, useEffect } from "react";
import Sidenav from "./sidenav";
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";
import background1 from "./../../assets/upload1_.png";

function Daily_Planner() {
  const ProjectIp = "http://localhost:5000";
  const [fileSelected, setFileSelected] = useState(false);
  const [block_data, setBlockdata] = useState([]);
  const [fixed_data, setFixeddata] = useState([]);
  const [selectedOption, setSelectedOption] = useState("default");
  const [subOptions, setSubOptions] = useState([]);
  const [selectedOption2, setSelectedOption2] = useState("default");
  const [subOptions2, setSubOptions2] = useState([]);
  const [subOption1, setSubOption1] = useState("");
  const [subOption2, setSubOption2] = useState("");
  const [selectedOption_fixed, setSelectedOption_fixed] = useState("default");
  const [subOptions_fixed, setSubOptions_fixed] = useState([]);
  const [selectedOption2_fixed, setSelectedOption2_fixed] = useState("default");
  const [subOptions2_fixed, setSubOptions2_fixed] = useState([]);
  const [subOption1_fixed, setSubOption1_fixed] = useState("");
  const [subOption2_fixed, setSubOption2_fixed] = useState("");
  const [commodity_fixed, setCommodity_fixed] = useState("");
  const [value_fixed, setValue_fixed] = useState("");
  const [TEFD, set_TEFD] = useState("");
  const [Scenerio, set_Scenerio] = useState("");
  const [solutionSolved, setSolutionSolved] = useState(false);
  const [scn, setscn] = useState(false);
  const [uploadst, setuploadst] = useState(false);
  const [Total_result, set_Total_Result] = useState(null);
  const [Relevant_result, set_Relevant_Result] = useState(null);

  const handleFileChange = (event) => {
    setFileSelected(event.target.files.length > 0);
  };

  const handleUploadConfig = async () => {
    if (!fileSelected) {
      alert("Please Select The File First");
      return;
    }

    try {
      const files = document.getElementById("uploadFile").files;
      const formData = new FormData();
      formData.append("uploadFile", files[0]);

      const response = await fetch(ProjectIp + "/uploadDailyFile_S2", {
        method: "POST",
        credentials: "include",
        body: formData,
      });

      if (!response.ok) {
        throw new Error("Network response was not ok");
      }

      const jsonResponse = await response.json();

      if (jsonResponse.status === 1) {
        alert("File Uploaded");
      } else {
        console.log(jsonResponse);
        alert("Error uploading file");
      }
    } catch (error) {
      console.error("Error during file upload:", error);
      alert("An error occurred during file upload. Please try again later.");
    }
  };

  const handleUploadConfig1 = async () => {
    if (!fileSelected) {
      alert("Please Select The File First");
      return;
    }

    try {
      const files = document.getElementById("uploadFile").files;
      const formData = new FormData();
      formData.append("uploadFile", files[0]);

      const response = await fetch(ProjectIp + "/uploadDailyFile_S1", {
        method: "POST",
        credentials: "include",
        body: formData,
      });

      if (!response.ok) {
        throw new Error("Network response was not ok");
      }

      const jsonResponse = await response.json();

      if (jsonResponse.status === 1) {
        alert("File Uploaded");
      } else {
        console.log(jsonResponse);
        alert("Error uploading file");
      }
    } catch (error) {
      console.error("Error during file upload:", error);
      alert("An error occurred during file upload. Please try again later.");
    }
  };

  const handleSolve = async () => {
    document.getElementById("toggle").checked = true;
    alert("This action will take time, click OK to continue.");
    if (Scenerio == "Scenerio 2") {
      setscn(true);
      setuploadst(true);
    }

    const payload = {
      TEFD: TEFD,
      origin_state: selectedOption,
      org_rhcode: subOption1,
      destination_state: selectedOption2,
      dest_rhcode: subOption2,
      block_data: block_data,
      Scenerio: Scenerio,
      confirmed_data: fixed_data,
    };
    try {
      const response = await fetch(ProjectIp + "/Daily_Planner", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(payload),
      });

      if (response.ok) {
        alert("Solution Done!, Now you can download results");
        setSolutionSolved(true);
      } else {
        console.error("Failed to send inputs. Status code:", response.status);
      }
    } catch (error) {
      console.error("Error sending inputs:", error);
    }
    document.getElementById("toggle").checked = false;
  };

  const fetchReservationId_Total_result = () => {
    var form = new FormData();
    fetch(ProjectIp + "/read_Daily_Planner_S1", {
      method: "POST",
      credentials: "include",
      body: form,
    })
      .then((response) => response.json())
      .then((data) => {
        const fetched_Total_Result = data;
        set_Total_Result(fetched_Total_Result);
      })
      .catch((error) => {
        console.error("Error:", error);
      });
  };
  const fetchReservationId_Revelant_result = () => {
    var form = new FormData();
    fetch(ProjectIp + "/read_Daily_Planner_S2", {
      method: "POST",
      credentials: "include",
      body: form,
    })
      .then((response) => response.json())
      .then((data) => {
        const fetched_Relevant_Result = data;
        set_Relevant_Result(fetched_Relevant_Result);
      })
      .catch((error) => {
        console.error("Error:", error);
      });
  };

  const handleDropdownChange = async (e) => {
    const selectedValue = e.target.value;
    setSelectedOption(selectedValue);
    const response = await fetch("/data/Updated_railhead_list.xlsx");
    const arrayBuffer = await response.arrayBuffer();
    const data = new Uint8Array(arrayBuffer);

    const workbook = XLSX.read(data, { type: "array" });

    // Assuming the Excel file has only one sheet
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Parse the sheet data into JSON format
    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    let dropdownOptions = [{ value: "", label: "Please select Railhead" }];
    for (let i = 0; i < jsonData.length; i++) {
      if (jsonData[i][1] == selectedValue) {
        dropdownOptions.push({ value: jsonData[i][0], label: jsonData[i][0] });
      }
    }
    dropdownOptions.sort((a, b) => a.label.localeCompare(b.label));

    setSubOptions(dropdownOptions);
    console.log(jsonData[1][1], dropdownOptions, selectedValue);
    setSubOptions(dropdownOptions);
  };

  const handleDropdownChange2 = async (e) => {
    const selectedValue = e.target.value;
    setSelectedOption2(selectedValue);
    const response = await fetch("/data/Updated_railhead_list.xlsx");
    const arrayBuffer = await response.arrayBuffer();
    const data = new Uint8Array(arrayBuffer);

    const workbook = XLSX.read(data, { type: "array" });

    // Assuming the Excel file has only one sheet
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Parse the sheet data into JSON format
    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    let dropdownOptions = [{ value: "", label: "Please select Railhead" }];
    for (let i = 0; i < jsonData.length; i++) {
      if (jsonData[i][1] == selectedValue) {
        dropdownOptions.push({ value: jsonData[i][0], label: jsonData[i][0] });
      }
    }
    dropdownOptions.sort((a, b) => a.label.localeCompare(b.label));

    setSubOptions2(dropdownOptions);
    console.log(jsonData[1][1], dropdownOptions, selectedValue);
    setSubOptions2(dropdownOptions);
  };

  const handleSubDropdownChange1 = (e) => {
    setSubOption1(e.target.value);
  };

  const handleSubDropdownChange2 = (e) => {
    setSubOption2(e.target.value);
  };

  const handleDropdownChange_fixed = async (e) => {
    const selectedValue_fixed = e.target.value;
    setSelectedOption_fixed(selectedValue_fixed);
    const response = await fetch("/data/Updated_railhead_list.xlsx");
    const arrayBuffer = await response.arrayBuffer();
    const data = new Uint8Array(arrayBuffer);

    const workbook = XLSX.read(data, { type: "array" });

    // Assuming the Excel file has only one sheet
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Parse the sheet data into JSON format
    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    let dropdownOptions = [{ value: "", label: "Please select Railhead" }];
    for (let i = 0; i < jsonData.length; i++) {
      if (jsonData[i][1] == selectedValue_fixed) {
        dropdownOptions.push({ value: jsonData[i][0], label: jsonData[i][0] });
      }
    }
    dropdownOptions.sort((a, b) => a.label.localeCompare(b.label));

    setSubOptions_fixed(dropdownOptions);
    console.log(jsonData[1][1], dropdownOptions, selectedValue_fixed);
    setSubOptions_fixed(dropdownOptions);
  };

  const handleDropdownChange2_fixed = async (e) => {
    const selectedValue = e.target.value;
    setSelectedOption2_fixed(selectedValue);
    const response = await fetch("/data/Updated_railhead_list.xlsx");
    const arrayBuffer = await response.arrayBuffer();
    const data = new Uint8Array(arrayBuffer);

    const workbook = XLSX.read(data, { type: "array" });

    // Assuming the Excel file has only one sheet
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Parse the sheet data into JSON format
    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    let dropdownOptions = [{ value: "", label: "Please select Railhead" }];
    for (let i = 0; i < jsonData.length; i++) {
      if (jsonData[i][1] == selectedValue) {
        dropdownOptions.push({ value: jsonData[i][0], label: jsonData[i][0] });
      }
    }
    dropdownOptions.sort((a, b) => a.label.localeCompare(b.label));

    setSubOptions2_fixed(dropdownOptions);
    console.log(jsonData[1][1], dropdownOptions, selectedValue);
    setSubOptions2_fixed(dropdownOptions);
  };

  const handleSubDropdownChange1_fixed = (e) => {
    setSubOption1_fixed(e.target.value);
  };

  const handleSubDropdownChange2_fixed = (e) => {
    setSubOption2_fixed(e.target.value);
  };

  const handleDeleteRow = (e) => {
    console.log(e);
    let block_data_ = block_data.filter((item) => item["id"] !== e);
    setBlockdata(block_data_);
  };

  const addConstraint = () => {
    // console.log(selectedOption, subOption1, selectedOption2, subOption2);
    if (selectedOption && subOption1 && selectedOption2 && subOption2) {
      setBlockdata((data) => [
        ...data,
        {
          origin_state: selectedOption,
          origin_railhead: subOption1,
          destination_state: selectedOption2,
          destination_railhead: subOption2,
          id: Date.now(),
        },
      ]);
      console.log(block_data);
      setSelectedOption("default");
      setSelectedOption2("default");
      setSubOptions([]);
      setSubOptions2([]);
    }
  };

  const addConstraint_fixed = () => {
    if (
      selectedOption_fixed &&
      subOption1_fixed &&
      selectedOption2_fixed &&
      subOption2_fixed &&
      commodity_fixed &&
      value_fixed
    ) {
      setFixeddata((data) => [
        ...data,
        {
          origin_state: selectedOption_fixed,
          origin_railhead: subOption1_fixed,
          destination_state: selectedOption2_fixed,
          destination_railhead: subOption2_fixed,
          commodity: commodity_fixed,
          value: value_fixed,
          id: Date.now(),
        },
      ]);
      console.log(fixed_data);
      setSelectedOption_fixed("default");
      setSelectedOption2_fixed("default");
      setSubOptions_fixed([]);
      setSubOptions2_fixed([]);
    }
  };

  const exportToExcel1 = () => {
    fetchReservationId_Total_result();
    if (Total_result == null) {
      window.alert("Fetching Result, Please Wait");
    } else {
      const workbook = XLSX.utils.book_new();
      Object.entries(Total_result).forEach(([column, data]) => {
        const parsedData = JSON.parse(data);
        const worksheet = XLSX.utils.json_to_sheet(parsedData);
        XLSX.utils.book_append_sheet(workbook, worksheet, column);
      });
      const excelBuffer = XLSX.write(workbook, {
        type: "array",
        bookType: "xlsx",
      });
      const excelBlob = new Blob([excelBuffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      saveAs(excelBlob, "Daily_Movement_Scenerio1.xlsx");
    }
  };

  const exportToExcel2 = () => {
    fetchReservationId_Revelant_result();
    if (Relevant_result == null) {
      window.alert("Fetching Result, Please Wait");
    } else {
      const workbook = XLSX.utils.book_new();
      Object.entries(Relevant_result).forEach(([column, data]) => {
        const parsedData = JSON.parse(data);
        const worksheet = XLSX.utils.json_to_sheet(parsedData);
        XLSX.utils.book_append_sheet(workbook, worksheet, column);
      });
      const excelBuffer = XLSX.write(workbook, {
        type: "array",
        bookType: "xlsx",
      });
      const excelBlob = new Blob([excelBuffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      saveAs(excelBlob, "Daily_Movement_results_Scenerio2.xlsx");
    }
  };

  return (
    <div className="page-container" style={{ backgroundColor: "#ebab44b0" }}>
      <Sidenav />
      <div
        className="page-content"
        style={{ backgroundImage: "url('static/img/bg8.jpg')" }}
      >
        <ul
          className="x-navigation x-navigation-horizontal x-navigation-panel"
          style={{ backgroundColor: "rgba(235, 171, 68, 0.69)" }}
        >
          <li className="xn-icon-button">
            <a href="#" className="x-navigation-minimize">
              <span className="fa fa-dedent" />
            </a>
          </li>
          <li
            className="xn-logo"
            style={{
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              width: "90%",
            }}
          >
            <span style={{ color: "black", fontSize: "32px" }}>
              Optimized Daily Plan
            </span>
            <a className="x-navigation-control"></a>
          </li>
        </ul>

        <ul className="breadcrumb">
          <li>
            <a href="/home">Home</a>
          </li>
          <li className="active">Daily plan</li>
        </ul>

        <div className="page-content-wrap">
          <div className="row">
            <div className="col-md-12">
              <br />
              <div className="row" style={{ marginLeft: "15px" }}>
                <div style={{ fontSize: "20px", fontWeight: "700" }}>
                  <i className="fa fa-file-excel-o" aria-hidden="true"></i>{" "}
                  Template
                </div>
                <form
                  action=""
                  encType="multipart/form-data"
                  id="uploadForm"
                  className="form-horizontal"
                >
                  <div
                    className="col-md-6"
                    style={{ marginTop: "15px", marginLeft: "50px" }}
                  >
                    <div className="form-group">
                      <div className="col-md-9">
                        <div className="input-group">
                          <span
                            className="input-group-addon"
                            style={{
                              backgroundColor: "rgba(235, 171, 68, 0.69)",
                            }}
                          >
                            <span className="fa fa-info" />
                          </span>

                          <input
                            type="file"
                            className="form-control"
                            onChange={handleFileChange}
                            id="uploadFile"
                            name="uploadFile"
                            defaultValue=""
                            required=""
                          />
                        </div>
                        <span className="help-block" style={{ color: "black" }}>
                          Choose Data Template
                        </span>
                      </div>
                    </div>
                  </div>
                  <div className="col-md-3">
                    {setuploadst && (
                      <div>
                        <img
                          className="upload_class"
                          src={background1}
                          id="uploadConfig"
                          onClick={handleUploadConfig}
                          disabled={!fileSelected}
                        />
                      </div>
                    )}
                    {!setuploadst && (
                      <div>
                        <img
                          className="upload_class"
                          src={background1}
                          id="uploadConfig"
                          onClick={handleUploadConfig1}
                          disabled={!fileSelected}
                        />
                      </div>
                    )}
                    <div style={{ marginTop: "-25px" }}>Click here</div>
                  </div>
                </form>
              </div>
              <br />
              <br />
              <div style={{ marginLeft: "15px" }}>
                <div style={{ fontSize: "20px", fontWeight: "700" }}>
                  <i className="fa fa-info-circle" aria-hidden="true"></i>{" "}
                  Configurations
                </div>
                <br />
                <form style={{ marginLeft: "50px" }}>
                  <label>
                    <strong
                      style={{
                        fontSize: "20px",
                        marginLeft: "15px",
                        color: "#9d0921",
                      }}
                    >
                      Select Matrix System
                    </strong>
                    <select
                      value={TEFD}
                      onChange={(e) => set_TEFD(e.target.value)}
                      style={{ marginLeft: "547px" }}
                    >
                      <option value="NON-TEFD">Non-TEFD</option>
                      <option value="TEFD">TEFD</option>
                    </select>
                  </label>
                  <label>
                    <strong
                      style={{
                        fontSize: "20px",
                        marginLeft: "15px",
                        color: "#9d0921",
                      }}
                    >
                      Select Scenerio
                    </strong>
                    <select
                      value={Scenerio}
                      onChange={(e) => set_Scenerio(e.target.value)}
                      style={{ marginLeft: "600px" }}
                    >
                      <option value="Scenerio 1">Scenerio 1</option>
                      <option value="Scenerio 2">Scenerio 2</option>
                    </select>
                  </label>
                  <br />
                  <br />

                  <p style={{ margin: 0, padding: 0 }}>
                    <strong
                      style={{
                        color: "#9d0921",
                        fontSize: "20px",
                        marginLeft: "15px",
                      }}
                    >
                      For Route Blocking:
                    </strong>
                  </p>
                  <br />
                  <div style={{ display: "flex", marginLeft: "20px" }}>
                    {/* <label htmlFor="origin_state"> */}
                    <div>
                      <strong style={{ fontSize: "16px", padding: "5px" }}>
                        Select Origin State
                      </strong>
                      <select
                        style={{ width: "200px", padding: "5px" }}
                        id="origin_state"
                        onChange={handleDropdownChange}
                        value={selectedOption}
                      >
                        <option value="default">Select Origin State</option>
                        <option value="Andhra Pradesh">Andhra Pradesh</option>
                        <option value="Bihar">Bihar</option>
                        <option value="Chattisgarh">Chattisgarh</option>
                        <option value="Goa">Goa</option>
                        <option value="Gujarat">Gujarat</option>
                        <option value="Haryana">Haryana</option>
                        <option value="Jammu & Kashmir">Jammu & Kashmir</option>
                        <option value="Jharkhand">Jharkhand</option>
                        <option value="Karnataka">Karnataka</option>
                        <option value="Kerala">Kerala</option>
                        <option value="MP">Madhya Pradesh</option>
                        <option value="Maharashtra">Maharashtra</option>
                        <option value="NE">North East</option>
                        <option value="Odisha">Odisha</option>
                        <option value="Punjab">Punjab</option>
                        <option value="Rajasthan">Rajasthan</option>
                        <option value="Tamil Nadu">Tamil Nadu</option>
                        <option value="Telangana">Telangana</option>
                        <option value="UP">Uttar Pradesh</option>
                        <option value="Uttarakhand">Uttarakhand</option>
                        <option value="West Bengal">West Bengal</option>
                      </select>
                    </div>

                    <div>
                      <strong style={{ fontSize: "16px", padding: "5px" }}>
                        Select Origin Railhead
                      </strong>
                      <select
                        id="origin_railhead"
                        style={{ width: "200px", padding: "5px" }}
                        onChange={handleSubDropdownChange1}
                        value={subOption1}
                      >
                        {subOptions.map((option) => (
                          <option key={option.value} value={option.value}>
                            {option.label}
                          </option>
                        ))}
                      </select>
                    </div>
                    {/* </label> */}
                    <div>
                      {/* <label htmlFor="deficit_state"> */}
                      <strong style={{ fontSize: "16px", padding: "5px" }}>
                        Select Destination State
                      </strong>
                      <select
                        style={{ width: "200px", padding: "5px" }}
                        id="deficit_state"
                        onChange={handleDropdownChange2}
                        value={selectedOption2}
                      >
                        <option value="default">
                          Select Destination State
                        </option>
                        <option value="Andhra Pradesh">Andhra Pradesh</option>
                        <option value="Bihar">Bihar</option>
                        <option value="Chattisgarh">Chattisgarh</option>
                        <option value="Goa">Goa</option>
                        <option value="Gujarat">Gujarat</option>
                        <option value="Haryana">Haryana</option>
                        <option value="Jammu & Kashmir">Jammu & Kashmir</option>
                        <option value="Jharkhand">Jharkhand</option>
                        <option value="Karnataka">Karnataka</option>
                        <option value="Kerala">Kerala</option>
                        <option value="MP">Madhya Pradesh</option>
                        <option value="Maharashtra">Maharashtra</option>
                        <option value="NE">North East</option>
                        <option value="Odisha">Odisha</option>
                        <option value="Punjab">Punjab</option>
                        <option value="Rajasthan">Rajasthan</option>
                        <option value="Tamil Nadu">Tamil Nadu</option>
                        <option value="Telangana">Telangana</option>
                        <option value="UP">Uttar Pradesh</option>
                        <option value="Uttarakhand">Uttarakhand</option>
                        <option value="West Bengal">West Bengal</option>
                      </select>
                    </div>
                    {/* </label> */}
                    {/* <label htmlFor="deficit_railhead"> */}
                    <div>
                      <strong style={{ fontSize: "16px", padding: "5px" }}>
                        Select Destination Railhead
                      </strong>
                      <select
                        id="deficit_railhead"
                        style={{ width: "200px", padding: "5px" }}
                        onChange={handleSubDropdownChange2}
                        value={subOption2}
                      >
                        {subOptions2.map((option) => (
                          <option key={option.value} value={option.value}>
                            {option.label}
                          </option>
                        ))}
                      </select>
                    </div>
                    {/* </label> */}
                    <div
                      style={{
                        padding: "5px",
                        margin: "2px",
                        float: "right",
                        width: "100px",
                        background: "orange",
                        padding: "auto",
                        cursor: "pointer",
                      }}
                      onClick={addConstraint}
                    >
                      <p style={{ textAlign: "center", marginTop: "20px" }}>
                        Add
                      </p>
                    </div>
                  </div>
                  <br />
                  <br />
                  <p style={{ margin: 0, padding: 0 }}>
                    <strong
                      style={{
                        color: "#9d0921",
                        fontSize: "20px",
                        marginLeft: "15px",
                      }}
                    >
                      For Route Fixing:
                    </strong>
                  </p>
                  <br />
                  <div style={{ display: "flex", marginLeft: "20px" }}>
                    {/* <label htmlFor="origin_state"> */}
                    <div>
                      <strong style={{ fontSize: "16px", padding: "5px" }}>
                        Select Origin State
                      </strong>
                      <select
                        style={{ width: "200px", padding: "5px" }}
                        id="origin_state"
                        onChange={handleDropdownChange_fixed}
                        value={selectedOption_fixed}
                      >
                        <option value="default">Select Origin State</option>
                        <option value="Andhra Pradesh">Andhra Pradesh</option>
                        <option value="Bihar">Bihar</option>
                        <option value="Chattisgarh">Chattisgarh</option>
                        <option value="Goa">Goa</option>
                        <option value="Gujarat">Gujarat</option>
                        <option value="Haryana">Haryana</option>
                        <option value="Jammu & Kashmir">Jammu & Kashmir</option>
                        <option value="Jharkhand">Jharkhand</option>
                        <option value="Karnataka">Karnataka</option>
                        <option value="Kerala">Kerala</option>
                        <option value="MP">Madhya Pradesh</option>
                        <option value="Maharashtra">Maharashtra</option>
                        <option value="NE">North East</option>
                        <option value="Odisha">Odisha</option>
                        <option value="Punjab">Punjab</option>
                        <option value="Rajasthan">Rajasthan</option>
                        <option value="Tamil Nadu">Tamil Nadu</option>
                        <option value="Telangana">Telangana</option>
                        <option value="UP">Uttar Pradesh</option>
                        <option value="Uttarakhand">Uttarakhand</option>
                        <option value="West Bengal">West Bengal</option>
                      </select>
                    </div>
                    {/* </label> */}
                    {/* <label htmlFor="origin_railhead"> */}
                    <div>
                      <strong style={{ fontSize: "16px", padding: "5px" }}>
                        Select Origin Railhead
                      </strong>
                      <select
                        id="origin_railhead"
                        style={{ width: "200px", padding: "5px" }}
                        onChange={handleSubDropdownChange1_fixed}
                        value={subOption1_fixed}
                      >
                        {subOptions_fixed.map((option) => (
                          <option key={option.value} value={option.value}>
                            {option.label}
                          </option>
                        ))}
                      </select>
                    </div>
                    {/* </label> */}
                    <div>
                      {/* <label htmlFor="deficit_state"> */}
                      <strong style={{ fontSize: "16px", padding: "5px" }}>
                        Select Destination State
                      </strong>
                      <select
                        style={{ width: "200px", padding: "5px" }}
                        id="deficit_state"
                        onChange={handleDropdownChange2_fixed}
                        value={selectedOption2_fixed}
                      >
                        <option value="default">
                          Select Destination State
                        </option>
                        <option value="Andhra Pradesh">Andhra Pradesh</option>
                        <option value="Bihar">Bihar</option>
                        <option value="Chattisgarh">Chattisgarh</option>
                        <option value="Goa">Goa</option>
                        <option value="Gujarat">Gujarat</option>
                        <option value="Haryana">Haryana</option>
                        <option value="Jammu & Kashmir">Jammu & Kashmir</option>
                        <option value="Jharkhand">Jharkhand</option>
                        <option value="Karnataka">Karnataka</option>
                        <option value="Kerala">Kerala</option>
                        <option value="MP">Madhya Pradesh</option>
                        <option value="Maharashtra">Maharashtra</option>
                        <option value="NE">North East</option>
                        <option value="Odisha">Odisha</option>
                        <option value="Punjab">Punjab</option>
                        <option value="Rajasthan">Rajasthan</option>
                        <option value="Tamil Nadu">Tamil Nadu</option>
                        <option value="Telangana">Telangana</option>
                        <option value="UP">Uttar Pradesh</option>
                        <option value="Uttarakhand">Uttarakhand</option>
                        <option value="West Bengal">West Bengal</option>
                      </select>
                    </div>
                    {/* </label> */}
                    {/* <label htmlFor="deficit_railhead"> */}
                    <div>
                      <strong style={{ fontSize: "16px", padding: "5px" }}>
                        Select Destination Railhead
                      </strong>
                      <select
                        id="deficit_railhead"
                        style={{ width: "200px", padding: "5px" }}
                        onChange={handleSubDropdownChange2_fixed}
                        value={subOption2_fixed}
                      >
                        {subOptions2_fixed.map((option) => (
                          <option key={option.value} value={option.value}>
                            {option.label}
                          </option>
                        ))}
                      </select>
                    </div>
                    {/* </label> */}
                    <div
                      style={{
                        padding: "5px",
                        margin: "2px",
                        float: "right",
                        width: "100px",
                        background: "orange",
                        padding: "auto",
                        cursor: "pointer",
                        marginTop: "50px",
                      }}
                      onClick={addConstraint_fixed}
                    >
                      <p style={{ textAlign: "center", marginTop: "20px" }}>
                        Add
                      </p>
                    </div>
                  </div>
                  <div
                    style={{
                      display: "flex",
                      marginLeft: "20px",
                      marginTop: "-20px",
                    }}
                  >
                    <strong
                      
                    >
                      Select Commodity
                    </strong>
                    <select
                      value={commodity_fixed}
                      onChange={(e) => setCommodity_fixed(e.target.value)}
                      style={{ 
                        marginLeft: "40px",
                      width : "200px",
                    padding : "5px" }}
                    >
                      <option value="RICE">Rice</option>
                      <option value="WHEAT">Wheat</option>
                    </select>
                    <div style={{ marginLeft: "50px" }}>
                      <strong style={{ fontSize: "16px", padding: "5px" }}>
                        Enter Value
                      </strong>
                      <input
                        type="number"
                        value={value_fixed}
                        onChange={(e) => setValue_fixed(e.target.value)}
                        style={{
                          marginLeft: "40px",
                          width: "200px",
                          padding: "5px",
                        }}
                      />
                    </div>
                  </div>
                  <br />
                  <br />
                </form>

                <div style={{ fontSize: "20px", fontWeight: "700" }}>
                  <i className="fa fa-list-alt" aria-hidden="true"></i> Optimal
                  Plan
                </div>
                <div
                  className="wrap__toggle"
                  style={{
                    textAlign: "center",
                    borderStyle: "solid",
                    borderColor: "#ebab44b0",
                  }}
                  onClick={handleSolve}
                >
                  <div className="wrap__toggle--bluetooth">
                    <span style={{ textAlign: "center", fontWeight: "bold" }}>
                      Generate Optimized Plan
                    </span>
                  </div>
                  <div className="wrap__toggle--toggler">
                    <label htmlFor="toggle">
                      <input
                        type="checkbox"
                        className="checkBox"
                        id="toggle"
                        onClick={handleSolve}
                      />
                      <span></span>
                    </label>
                  </div>
                </div>
                <br />
                {/* <div>
              <br/>
                    <DynamicTable/>
                  </div> */}
                <br />
                {solutionSolved && (
                  <div>
                    {scn && (
                      <div>
                        <button
                          style={{ color: "white", marginLeft: "15px" }}
                          className="btn btn-danger dropdown-toggle"
                          onClick={() => exportToExcel2()}
                        >
                          <i className="fa fa-bars"></i> Download
                          Railhead-Railhead Detailed Plan
                        </button>
                      </div>
                    )}
                    {!scn && (
                      <div>
                        <button
                          style={{ color: "white", marginLeft: "15px" }}
                          className="btn btn-danger dropdown-toggle"
                          onClick={() => exportToExcel1()}
                        >
                          <i className="fa fa-bars"></i> Download
                          Railhead-Railhead Detailed Plan
                        </button>
                      </div>
                    )}
                  </div>
                )}
                <br />
                <br />
                {!solutionSolved && (
                  <div>
                    <div
                      style={{
                        fontSize: "20px",
                        fontWeight: "700",
                        textAlign: "center",
                      }}
                    >
                      Route Block
                    </div>
                    <table>
                      <thead>
                        <tr style={{ margin: "auto" }}>
                          <th style={{ padding: "10px", width: "15%" }}>
                            Origin State
                          </th>
                          <th style={{ padding: "10px", width: "15%" }}>
                            Origin Railhead
                          </th>
                          <th style={{ padding: "10px", width: "15%" }}>
                            Destination State
                          </th>
                          <th style={{ padding: "10px", width: "15%" }}>
                            Destination Railhead
                          </th>
                          <th style={{ padding: "10px", width: "15%" }}>
                            Delete
                          </th>
                        </tr>
                        {/* <tr  style={{ padding: "10px", width: "100%" , textAlign:'center'}}>
                      <div style={{textAlign:'center', width:'100%'}}>Routes Block</div></tr> */}
                      </thead>
                      <tbody>
                        {/* <tr style={{ margin: "auto" }}>
                      <th style={{ padding: "10px", width: "15%" }}>
                        Origin State
                      </th>
                      <th style={{ padding: "10px", width: "15%" }}>
                        Origin Railhead
                      </th>
                      <th style={{ padding: "10px", width: "15%" }}>
                        Destination State
                      </th>
                      <th style={{ padding: "10px", width: "15%" }}>
                        Destination Railhead
                      </th>
                      <th style={{ padding: "10px", width: "15%" }}>Delete</th>
                    </tr> */}
                        {block_data.map((item) => (
                          <tr>
                            <td>{item.origin_state}</td>
                            <td>{item.origin_railhead}</td>
                            <td>{item.destination_state}</td>
                            <td>{item.destination_railhead}</td>
                            <td>
                              <span
                                style={{
                                  cursor: "pointer",
                                  color: "#ff0000",
                                  fontSize: "1.2rem",
                                }}
                                onClick={() => handleDeleteRow(item.id)}
                                title="Delete"
                              >
                                &times;
                              </span>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}

                <br />
                <br />
                {!solutionSolved && (
                  <div>
                    <div
                      style={{
                        fontSize: "20px",
                        fontWeight: "700",
                        textAlign: "center",
                      }}
                    >
                      Fixed Routes
                    </div>
                    <table>
                      <thead>
                        <tr style={{ margin: "auto" }}>
                          <th style={{ padding: "10px", width: "15%" }}>
                            Origin State
                          </th>
                          <th style={{ padding: "10px", width: "15%" }}>
                            Origin Railhead
                          </th>
                          <th style={{ padding: "10px", width: "15%" }}>
                            Destination State
                          </th>
                          <th style={{ padding: "10px", width: "15%" }}>
                            Destination Railhead
                          </th>
                          <th style={{ padding: "10px", width: "15%" }}>
                            Commodity
                          </th>
                          <th style={{ padding: "10px", width: "15%" }}>
                          </th>
                          <th style={{ padding: "10px", width: "15%" }}>
                            Delete
                          </th>
                        </tr>
                        {/* <tr  style={{ padding: "10px", width: "100%" , textAlign:'center'}}>
                      <div style={{textAlign:'center', width:'100%'}}>Routes Block</div></tr> */}
                      </thead>
                      <tbody>
                        {/* <tr style={{ margin: "auto" }}>
                      <th style={{ padding: "10px", width: "15%" }}>
                        Origin State
                      </th>
                      <th style={{ padding: "10px", width: "15%" }}>
                        Origin Railhead
                      </th>
                      <th style={{ padding: "10px", width: "15%" }}>
                        Destination State
                      </th>
                      <th style={{ padding: "10px", width: "15%" }}>
                        Destination Railhead
                      </th>
                      <th style={{ padding: "10px", width: "15%" }}>Delete</th>
                    </tr> */}
                        {fixed_data.map((item) => (
                          <tr>
                            <td>{item.origin_state}</td>
                            <td>{item.origin_railhead}</td>
                            <td>{item.destination_state}</td>
                            <td>{item.destination_railhead}</td>
                            <td>{item.commodity}</td>
                            <td>{item.value}</td>
                            <td>
                              <span
                                style={{
                                  cursor: "pointer",
                                  color: "#ff0000",
                                  fontSize: "1.2rem",
                                }}
                                onClick={() => handleDeleteRow(item.id)}
                                title="Delete"
                              >
                                &times;
                              </span>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            </div>
          </div>
          {/* {solutionSolved && (
            <div>
              <p style={{ display: "inline", marginLeft: "25px" }}>
                <strong style={{ fontSize: "16px" }}>
                  Optimal Cost of Transportation is INR{" "}
                  <span style={{ color: "#FF0509" }}>{cost}</span> Lakhs
                </strong>
              </p>
            </div>
          )} */}

          <br />
          <div className="panel-heading">
            <h3 className="panel-title"></h3>
            <div className="btn-group pull-left">
              <br />
              <br />
            </div>
          </div>
          <br />
          <br />
          {/* </div> */}
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
        </div>
      </div>
    </div>
  );
}

export default Daily_Planner;