import React, { useState } from "react";
import Sidenav from "./sidenav";
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";
import config from "../../config";
import style from "./Daily_Planner.css";

function Daily_Planner() {
  const ProjectIp = config.serverUrl;
  const [fileSelected, setFileSelected] = useState(false);
  const [inline_value_rice, setInlineValueRice] = useState("");
  const [inline_value_wheat, setInlineValueWheat] = useState("");
  const [block_data, setBlockdata] = useState([]);
  const [block_data2, setBlockdata2] = useState([]);
  const [block_dataWheat2, setBlockdataWheat2] = useState([]);
  const [block_data3, setBlockdata3] = useState([]);
  const [block_dataWheat3, setBlockdataWheat3] = useState([]);
  const [rice_destination, setRiceDestination] = useState([]);
  const [wheat_destination, setWheatDestination] = useState([]);
  const [fixed_data, setFixeddata] = useState([]);
  const [selectedOption, setSelectedOption] = useState("default");
  const [subOptions, setSubOptions] = useState([]);
  const [selectedOption2, setSelectedOption2] = useState("default");
  const [subOptions2, setSubOptions2] = useState([]);
  const [selectedOption3, setSelectedOption3] = useState("default");
  const [subOptions3, setSubOptions3] = useState([]);
  const [selectedOptionWheat3, setSelectedOptionWheat3] = useState("default");
  const [subOptionsWheat3, setSubOptionsWheat3] = useState([]);
  const [selectedOption4, setSelectedOption4] = useState("default");
  const [subOptions4, setSubOptions4] = useState([]);
  const [selectedOptionWheat4, setSelectedOptionWheat4] = useState("default");
  const [subOptionsWheat4, setSubOptionsWheat4] = useState([]);
  const [selectedOption5, setSelectedOption5] = useState("default");
  const [subOptions5, setSubOptions5] = useState([]);
  const [selectedOption6, setSelectedOption6] = useState("default");
  const [subOptions6, setSubOptions6] = useState([]);
  const [selectedOptionWheat5, setSelectedOptionWheat5] = useState("default");
  const [subOptionsWheat5, setSubOptionsWheat5] = useState([]);
  const [selectedOptionWheat6, setSelectedOptionWheat6] = useState("default");
  const [subOptionsWheat6, setSubOptionsWheat6] = useState([]);
  const [subOption1, setSubOption1] = useState("");
  const [subOption2, setSubOption2] = useState("");
  const [subOption3, setSubOption3] = useState("");
  const [subOptionWheat3, setSubOptionWheat3] = useState("");
  const [subOption4, setSubOption4] = useState("");
  const [subOptionWheat4, setSubOptionWheat4] = useState("");
  const [subOption5, setSubOption5] = useState("");
  const [subOption6, setSubOption6] = useState("");
  const [subOptionWheat5, setSubOptionWheat5] = useState("");
  const [subOptionWheat6, setSubOptionWheat6] = useState("");
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
  const [Total_result, set_Total_Result] = useState();
  const [Relevant_result, set_Relevant_Result] = useState(null);
  const [excelData, setExcelData] = useState({});
  const [activeSheetName, setActiveSheetName] = useState(null);
  const [sheet, setSheet] = useState(null);
  const [updateExcel, setUpdateExcel] = useState(false);
  const [updateExcel2, setUpdateExcel2] = useState(false);
  const [modifiedExcel, setModifiedExcel] = useState({});
  const [isLoading, setIsLoading] = useState(false);
  const [isLoading2, setIsLoading2] = useState(false);
  const [isLoading3, setIsLoading3] = useState(false);
  const [number_check1, setnumber_check1] = useState(0);
  const [number_check2, setnumber_check2] = useState(0);
  const [supplyWeatCount, setSupplyWeatCount] = useState(0);
  const [destinationWheatCount, setDestinationWheatCount] = useState(0);
  const [showMessage, setShowMessage] = useState(false);
  const [riceData, setRiceData] = useState(false);
  const [wheatData, setWheatData] = useState(false);
  const [downloadMessage, setDownloadMessage] = useState(false);
  const [progress, setProgress] = useState([]);
  const [filterRailHead, setfilterRailHead] = useState([]);
  // Define a Set to keep track of added railheads
  const [addedRailheads, setAddedRailheads] = useState(new Set());

  // Block_data for blocking, fixed_data for fixing, block_data3 for rice_origin, block_data4 for rice_destination
  const handleCellChange = (sheetName, rowIndex, columnIndex, newValue) => {
    const updatedData = { ...excelData };
    updatedData[sheetName][rowIndex][columnIndex] = newValue;
    setExcelData(updatedData);
  };

  const handleFileChange = (event) => {
    setFileSelected(event.target.files.length > 0);
    const files = document.getElementById("uploadFile").files;
    const reader = new FileReader();
    const file = files[0];
    reader.onload = async (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const sheetsData = {};
      workbook.SheetNames.forEach((sheetName) => {
        const worksheet = workbook.Sheets[sheetName];
        sheetsData[sheetName] = XLSX.utils.sheet_to_json(worksheet, {
          header: 1,
        });
      });

      setExcelData(sheetsData);
      setActiveSheetName(workbook.SheetNames[0]);
    };

    reader.readAsArrayBuffer(file);
  };

  const update_excel = async () => {
    setUpdateExcel2(false);
    const response = await fetch(ProjectIp + "/getDaily1ExcelData");
    const arrayBuffer = await response.arrayBuffer();
    const data = new Uint8Array(arrayBuffer);
    const workbook = XLSX.read(data, { type: "array" });
    const sheetsData = {};
    workbook.SheetNames.forEach((sheetName) => {
      const worksheet = workbook.Sheets[sheetName];
      sheetsData[sheetName] = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
      });
    });

    setExcelData(sheetsData);
    setActiveSheetName(workbook.SheetNames[0]);
    setUpdateExcel(true);
  };

  const update_excel2 = async () => {
    setUpdateExcel(false);
    const response = await fetch(ProjectIp + "/getDaily2ExcelData");
    const arrayBuffer = await response.arrayBuffer();
    const data = new Uint8Array(arrayBuffer);
    const workbook = XLSX.read(data, { type: "array" });
    const sheetsData = {};
    workbook.SheetNames.forEach((sheetName) => {
      const worksheet = workbook.Sheets[sheetName];
      sheetsData[sheetName] = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
      });
    });

    setExcelData(sheetsData);
    setActiveSheetName(workbook.SheetNames[0]);
    setUpdateExcel2(true);
  };

  const save_excel = async () => {
    const newWorkbook = XLSX.utils.book_new();
    Object.keys(excelData).forEach((sheetName) => {
      const worksheet = XLSX.utils.json_to_sheet(excelData[sheetName]);
      XLSX.utils.book_append_sheet(newWorkbook, worksheet, sheetName);
    });

    try {
      var scenario;
      if (updateExcel === true) {
        scenario = "/Modify_Daily_Template_S01";
      } else {
        scenario = "/Modify_Daily_Template_S02";
      }
      const response = await fetch(ProjectIp + scenario, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(newWorkbook),
      });
      if (response.ok) {
        console.log("Data sent to backend successfully");
      } else {
        console.error("Failed to send data to backend");
      }
    } catch (error) {
      console.error("Error sending data:", error);
    }
    setUpdateExcel(false);
    setUpdateExcel2(false);
    document.getElementById("console_").innerHTML +=
      "Template has been updated" + "<br/><br/>";
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

  console.log({ Total_result });

  const handleSolve = async () => {
    setDownloadMessage(false);
    if (
      number_check1 < number_check2 ||
      supplyWeatCount < destinationWheatCount
    ) {
      alert("Destination indents more than Supply indents Please check");
      setIsLoading(false);
      document.getElementById("toggle").checked = false;
      return;
    }
    if (isLoading) return;
    setIsLoading(true);

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
      rice_origin: block_data3,
      rice_destination: rice_destination,
      rice_inline: block_data2,
      rice_inline_value: inline_value_rice,
      wheat_origin: block_dataWheat3,
      wheat_destination: wheat_destination,
      wheat_inline: block_dataWheat2,
      wheat_inline_value: inline_value_wheat,
    };

    try {
      const response = await fetch(ProjectIp + "/Daily_Planner", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(payload),
      });
      fetchReservationId_Total_result();
      fetchReservationId_Revelant_result();
      if (response.ok) {
        setSolutionSolved(true);
      } else {
        console.error("Failed to send inputs. Status code:", response.status);
      }
    } catch (error) {
      console.error("Error sending inputs:", error);
    } finally {
      setIsLoading(false); // Reset loading state
      setDownloadMessage(true);
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
        set_Total_Result(data);
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
    let dropdownOptions = [];
    let dropdownOptions_default = {
      value: "",
      label: "Please select Railhead",
    };
    for (let i = 0; i < jsonData.length; i++) {
      if (
        jsonData[i][1] &&
        jsonData[i][1].trim().toLowerCase() ===
          selectedValue.trim().toLowerCase()
      ) {
        dropdownOptions.push({ value: jsonData[i][0], label: jsonData[i][0] });
      }
    }
    dropdownOptions.sort((a, b) => a.label.localeCompare(b.label));
    // dropdownOptions=dropdownOptions_default+dropdownOptions;
    dropdownOptions.unshift(dropdownOptions_default);
    setSubOptions(dropdownOptions);
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
    let dropdownOptions = [];
    let dropdownOptions_default = {
      value: "",
      label: "Please select Railhead",
    };
    for (let i = 0; i < jsonData.length; i++) {
      if (
        jsonData[i][1] &&
        jsonData[i][1].trim().toLowerCase() ===
          selectedValue.trim().toLowerCase()
      ) {
        dropdownOptions.push({ value: jsonData[i][0], label: jsonData[i][0] });
      }
    }
    dropdownOptions.sort((a, b) => a.label.localeCompare(b.label));
    // dropdownOptions=dropdownOptions_default+dropdownOptions;
    dropdownOptions.unshift(dropdownOptions_default);
    setSubOptions2(dropdownOptions);
  };

  const handleDropdownChange3 = async (e) => {
    const selectedValue = e.target.value;
    console.log({ selectedValue });
    setSelectedOption3(selectedValue);
    const response = await fetch("/data/Updated_railhead_list.xlsx");
    const arrayBuffer = await response.arrayBuffer();
    const data = new Uint8Array(arrayBuffer);

    const workbook = XLSX.read(data, { type: "array" });

    // Assuming the Excel file has only one sheet
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Parse the sheet data into JSON format
    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    let dropdownOptions = [];
    let dropdownOptions_default = {
      value: "",
      label: "Please select Railhead",
    };
    for (let i = 0; i < jsonData.length; i++) {
      if (
        jsonData[i][1] &&
        jsonData[i][1].trim().toLowerCase() ===
          selectedValue.trim().toLowerCase()
      ) {
        dropdownOptions.push({ value: jsonData[i][0], label: jsonData[i][0] });
      }
    }
    dropdownOptions.sort((a, b) => a.label.localeCompare(b.label));
    // dropdownOptions=dropdownOptions_default+dropdownOptions;
    dropdownOptions.unshift(dropdownOptions_default);
    setSubOptions3(dropdownOptions);
  };

  const handleDropdownChangeWheat3 = async (e) => {
    const selectedValue = e.target.value;
    setSelectedOptionWheat3(selectedValue);
    const response = await fetch("/data/Updated_railhead_list.xlsx");
    const arrayBuffer = await response.arrayBuffer();
    const data = new Uint8Array(arrayBuffer);

    const workbook = XLSX.read(data, { type: "array" });

    // Assuming the Excel file has only one sheet
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Parse the sheet data into JSON format
    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    let dropdownOptions = [];
    let dropdownOptions_default = {
      value: "",
      label: "Please select Railhead",
    };
    for (let i = 0; i < jsonData.length; i++) {
      if (
        jsonData[i][1] &&
        jsonData[i][1].trim().toLowerCase() ===
          selectedValue.trim().toLowerCase()
      ) {
        dropdownOptions.push({ value: jsonData[i][0], label: jsonData[i][0] });
      }
    }
    dropdownOptions.sort((a, b) => a.label.localeCompare(b.label));
    // dropdownOptions=dropdownOptions_default+dropdownOptions;
    dropdownOptions.unshift(dropdownOptions_default);
    setSubOptionsWheat3(dropdownOptions);
  };

  const handleDropdownChange5 = async (e) => {
    const selectedValue = e.target.value;
    console.log(selectedValue);
    setSelectedOption5(selectedValue);
    const response = await fetch("/data/Updated_railhead_list.xlsx");
    const arrayBuffer = await response.arrayBuffer();
    const data = new Uint8Array(arrayBuffer);

    const workbook = XLSX.read(data, { type: "array" });

    // Assuming the Excel file has only one sheet
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Parse the sheet data into JSON format
    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    let dropdownOptions = [];
    let dropdownOptions_default = {
      value: "",
      label: "Please select Railhead",
    };
    for (let i = 0; i < jsonData.length; i++) {
      if (jsonData[i][1] === selectedValue) {
        dropdownOptions.push({ value: jsonData[i][0], label: jsonData[i][0] });
      }
    }
    dropdownOptions.sort((a, b) => a.label.localeCompare(b.label));
    // dropdownOptions=dropdownOptions_default+dropdownOptions;
    dropdownOptions.unshift(dropdownOptions_default);
    setSubOptions5(dropdownOptions);
  };

  const handleDropdownChangeWheat5 = async (e) => {
    const selectedValue = e.target.value;
    setSelectedOptionWheat5(selectedValue);
    const response = await fetch("/data/Updated_railhead_list.xlsx");
    const arrayBuffer = await response.arrayBuffer();
    const data = new Uint8Array(arrayBuffer);

    const workbook = XLSX.read(data, { type: "array" });

    // Assuming the Excel file has only one sheet
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Parse the sheet data into JSON format
    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    let dropdownOptions = [];
    let dropdownOptions_default = {
      value: "",
      label: "Please select Railhead",
    };
    for (let i = 0; i < jsonData.length; i++) {
      if (jsonData[i][1] === selectedValue) {
        dropdownOptions.push({ value: jsonData[i][0], label: jsonData[i][0] });
      }
    }
    dropdownOptions.sort((a, b) => a.label.localeCompare(b.label));
    // dropdownOptions=dropdownOptions_default+dropdownOptions;
    dropdownOptions.unshift(dropdownOptions_default);
    setSubOptionsWheat5(dropdownOptions);
  };

  const handleDropdownChange6 = async (e) => {
    const selectedValue = e.target.value;
    setSelectedOption6(selectedValue);

    // Check if the railhead for the selected state is already added

    const response = await fetch("/data/Updated_railhead_list.xlsx");
    const arrayBuffer = await response.arrayBuffer();
    const data = new Uint8Array(arrayBuffer);

    const workbook = XLSX.read(data, { type: "array" });

    // Assuming the Excel file has only one sheet
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Parse the sheet data into JSON format
    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    let dropdownOptions = [];
    let dropdownOptions_default = {
      value: "",
      label: "Please select Railhead",
    };
    for (let i = 0; i < jsonData.length; i++) {
      if (jsonData[i][1] == selectedValue) {
        dropdownOptions.push({
          value: jsonData[i][0],
          label: jsonData[i][0],
        });
      }
    }
    dropdownOptions.sort((a, b) => a.label.localeCompare(b.label));
    dropdownOptions.unshift(dropdownOptions_default);
    setSubOptions6(dropdownOptions);

    // Add the railhead to the set to indicate it's added for this state
    setAddedRailheads(new Set(addedRailheads).add(selectedValue));
  };

  const handleDropdownChangeWheat6 = async (e) => {
    const selectedValue = e.target.value;
    setSelectedOptionWheat6(selectedValue);
    const response = await fetch("/data/Updated_railhead_list.xlsx");
    const arrayBuffer = await response.arrayBuffer();
    const data = new Uint8Array(arrayBuffer);

    const workbook = XLSX.read(data, { type: "array" });

    // Assuming the Excel file has only one sheet
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Parse the sheet data into JSON format
    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    let dropdownOptions = [];
    let dropdownOptions_default = {
      value: "",
      label: "Please select Railhead",
    };
    for (let i = 0; i < jsonData.length; i++) {
      if (jsonData[i][1] == selectedValue) {
        dropdownOptions.push({ value: jsonData[i][0], label: jsonData[i][0] });
      }
    }
    dropdownOptions.sort((a, b) => a.label.localeCompare(b.label));
    // dropdownOptions=dropdownOptions_default+dropdownOptions;
    dropdownOptions.unshift(dropdownOptions_default);
    setSubOptionsWheat6(dropdownOptions);
  };

  const handleDropdownChange4 = async (e) => {
    const selectedValue = e.target.value;
    setSelectedOption4(selectedValue);
    const response = await fetch("/data/Updated_railhead_list.xlsx");
    const arrayBuffer = await response.arrayBuffer();
    const data = new Uint8Array(arrayBuffer);

    const workbook = XLSX.read(data, { type: "array" });

    // Assuming the Excel file has only one sheet
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Parse the sheet data into JSON format
    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    let dropdownOptions = [];
    let dropdownOptions_default = {
      value: "",
      label: "Please select Railhead",
    };
    for (let i = 0; i < jsonData.length; i++) {
      if (
        jsonData[i][1] &&
        jsonData[i][1].trim().toLowerCase() ===
          selectedValue.trim().toLowerCase()
      ) {
        dropdownOptions.push({ value: jsonData[i][0], label: jsonData[i][0] });
      }
    }
    dropdownOptions.sort((a, b) => a.label.localeCompare(b.label));
    // dropdownOptions=dropdownOptions_default+dropdownOptions;
    dropdownOptions.unshift(dropdownOptions_default);
    setSubOptions4(dropdownOptions);
  };

  const handleDropdownChangeWheat4 = async (e) => {
    const selectedValue = e.target.value;
    setSelectedOptionWheat4(selectedValue);
    const response = await fetch("/data/Updated_railhead_list.xlsx");
    const arrayBuffer = await response.arrayBuffer();
    const data = new Uint8Array(arrayBuffer);

    const workbook = XLSX.read(data, { type: "array" });

    // Assuming the Excel file has only one sheet
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Parse the sheet data into JSON format
    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    let dropdownOptions = [];
    let dropdownOptions_default = {
      value: "",
      label: "Please select Railhead",
    };
    for (let i = 0; i < jsonData.length; i++) {
      if (
        jsonData[i][1] &&
        jsonData[i][1].trim().toLowerCase() ===
          selectedValue.trim().toLowerCase()
      ) {
        dropdownOptions.push({ value: jsonData[i][0], label: jsonData[i][0] });
      }
    }
    dropdownOptions.sort((a, b) => a.label.localeCompare(b.label));
    // dropdownOptions=dropdownOptions_default+dropdownOptions;
    dropdownOptions.unshift(dropdownOptions_default);
    setSubOptionsWheat4(dropdownOptions);
  };

  const handleSubDropdownChange1 = (e) => {
    setSubOption1(e.target.value);
  };

  const handleSubDropdownChange2 = (e) => {
    setSubOption2(e.target.value);
  };
  const handleSubDropdownChange3 = (e) => {
    setSubOption3(e.target.value);
  };
  const handleSubDropdownChangeWheat3 = (e) => {
    setSubOptionWheat3(e.target.value);
  };

  const handleSubDropdownChange4 = (e) => {
    setSubOption4(e.target.value);
  };
  const handleSubDropdownChangeWheat4 = (e) => {
    setSubOptionWheat4(e.target.value);
  };

  const handleSubDropdownChange5 = (e) => {
    setSubOption5(e.target.value);
  };

  const handleSubDropdownChangeWheat5 = (e) => {
    setSubOptionWheat5(e.target.value);
  };

  const handleSubDropdownChange6 = (e) => {
    setSubOption6(e.target.value);
  };
  const handleSubDropdownChangeWheat6 = (e) => {
    setSubOptionWheat6(e.target.value);
  };

  const handleInlineValueWheat = (e) => {
    setInlineValueWheat(e.target.value);
  };
  const handleInlineValueRice = (e) => {
    setInlineValueRice(e.target.value);
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
    let dropdownOptions = [];
    let dropdownOptions_default = {
      value: "",
      label: "Please select Railhead",
    };
    for (let i = 0; i < jsonData.length; i++) {
      if (
        jsonData[i][1] &&
        jsonData[i][1].trim().toLowerCase() ===
          selectedValue_fixed.trim().toLowerCase()
      ) {
        dropdownOptions.push({ value: jsonData[i][0], label: jsonData[i][0] });
      }
    }
    dropdownOptions.sort((a, b) => a.label.localeCompare(b.label));
    // dropdownOptions=dropdownOptions_default+dropdownOptions;
    dropdownOptions.unshift(dropdownOptions_default);
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
    let dropdownOptions = [];
    let dropdownOptions_default = {
      value: "",
      label: "Please select Railhead",
    };
    for (let i = 0; i < jsonData.length; i++) {
      if (
        jsonData[i][1] &&
        jsonData[i][1].trim().toLowerCase() ===
          selectedValue.trim().toLowerCase()
      ) {
        dropdownOptions.push({ value: jsonData[i][0], label: jsonData[i][0] });
      }
    }
    dropdownOptions.sort((a, b) => a.label.localeCompare(b.label));
    // dropdownOptions=dropdownOptions_default+dropdownOptions;
    dropdownOptions.unshift(dropdownOptions_default);
    setSubOptions2_fixed(dropdownOptions);
  };

  const handleSubDropdownChange1_fixed = (e) => {
    setSubOption1_fixed(e.target.value);
  };

  const handleSubDropdownChange2_fixed = (e) => {
    setSubOption2_fixed(e.target.value);
  };

  const handleDeleteRow = (e) => {
    let block_data_ = block_data.filter((item) => item["id"] !== e);
    setBlockdata(block_data_);
    setDownloadMessage(false);
  };
  const handleDeleteRow_fixed = (e) => {
    let fixed_data_ = fixed_data.filter((item) => item["id"] !== e);
    setFixeddata(fixed_data_);
    setDownloadMessage(false);
  };
  const handleDeleteRow_inline = (e) => {
    let fixed_data_ = block_data2.filter((item) => item["id"] !== e);
    setBlockdata2(fixed_data_);
    setDownloadMessage(false);
  };
  const handleDeleteRow_inlineWheat = (e) => {
    let fixed_data_ = block_dataWheat2.filter((item) => item["id"] !== e);
    setBlockdataWheat2(fixed_data_);
    setDownloadMessage(false);
  };
  const handleDeleteRow_Rice_s = (e) => {
    let block_data3_ = block_data3.filter((item) => item["id"] !== e);
    setBlockdata3(block_data3_);
    setnumber_check1(number_check1 - 1);
    setDownloadMessage(false);
  };

  const handleDeleteRow_Wheat_s = (e) => {
    let block_data3_ = block_dataWheat3.filter((item) => item["id"] !== e);
    setBlockdataWheat3(block_data3_);
    setSupplyWeatCount(supplyWeatCount - 1);
    setDownloadMessage(false);
  };

  const handleDeleteRow_Rice__dest = (index) => {
    let rice_destination_ = rice_destination.filter((item, i) => i !== index);
    setRiceDestination(rice_destination_);
    setnumber_check2(number_check2 - 1);
    setDownloadMessage(false);
  };

  const handleDeleteRow_Wheat__dest = (index) => {
    let wheat_destination_ = wheat_destination.filter((item, i) => i !== index);
    setWheatDestination(wheat_destination_);
    setDestinationWheatCount(destinationWheatCount - 1);
    setDownloadMessage(false);
  };

  const addConstraint = (e) => {
    e.preventDefault();
    if (selectedOption && subOption1 && selectedOption2 && subOption2) {
      // Check if origin and destination railheads are the same
      // if (subOption1 === subOption2) {
      //   alert("Origin and destination railheads cannot be the same.");
      //   return; // Do not proceed further
      // }

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

      // Reset options and suboptions
      setSelectedOption("default");
      setSelectedOption2("default");
      setSubOptions([]);
      setSubOptions2([]);

      setProgress((prev) => [
        ...prev,
        `Route from ${subOption1} to ${subOption2} has been blocked`,
      ]);
    }
  };

  const addConstraint2 = async (e) => {
    e.preventDefault();

    // Check if necessary options are selected
    if (selectedOption5 && subOption5 && selectedOption6 && subOption6) {
      // Check if the origin and destination railheads are the same
      // if (selectedOption5 === selectedOption6 && subOption5 === subOption6) {
      //   alert("Origin and destination railheads cannot be the same.");
      //   return; // Do not proceed further
      // }

      // Update block data with the new constraint
      setBlockdata2((data) => [
        ...data,
        {
          origin_state: selectedOption5,
          origin_railhead: subOption5,
          destination_state: selectedOption6,
          destination_railhead: subOption6,
          id: Date.now(),
        },
      ]);

      // Reset options and suboptions for rice
      setSelectedOption5("default");
      setSelectedOption6("default");
      setSubOptions5([]);
      setSubOptions6([]);

      if (isLoading2) return; // Prevent additional clicks while loading
      setIsLoading2(true);

      try {
        const payload1 = {
          rice_inline: block_data2,
          rice_inline_value: inline_value_rice,
          wheat_inline: block_dataWheat2,
          wheat_inline_value: inline_value_wheat,
        };

        const response2 = await fetch(ProjectIp + "/Daily_Planner_Check", {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
          },
          body: JSON.stringify(payload1),
        });

        const responseData1 = await response2.json(); // Parse response JSON

        if (responseData1.status === "NO") {
          alert("Distance is not within range. Please check again.");
        }
      } catch (error) {
        console.error("Error sending inputs:", error);
      } finally {
        // Reset loading state
        setIsLoading2(false);
      }
    }

    // Add progress message
    setProgress((prev) => [
      ...prev,
      "New Inline details have been added for rice",
    ]);
  };

  const addConstraintWheat2 = async (e) => {
    e.preventDefault();

    // Check if necessary options are selected
    if (
      selectedOptionWheat5 &&
      subOptionWheat5 &&
      selectedOptionWheat6 &&
      subOptionWheat6
    ) {
      // Check if origin and destination railheads for wheat are the same
      // if (subOptionWheat5 === subOptionWheat6) {
      //   alert("Origin and destination railheads for wheat cannot be the same.");
      //   return; // Do not proceed further
      // }

      // Update wheat block data with the new constraint
      let data = [
        {
          origin_state: selectedOptionWheat5,
          origin_railhead: subOptionWheat5,
          destination_state: selectedOptionWheat6,
          destination_railhead: subOptionWheat6,
          id: Date.now(),
        },
      ];

      setBlockdataWheat2((prevData) => [...prevData, ...data]);

      // Reset options and suboptions for wheat
      setSelectedOptionWheat5("default");
      setSelectedOptionWheat6("default");
      setSubOptionWheat5([]);
      setSubOptionWheat6([]);

      if (isLoading3) return; // Prevent additional clicks while loading
      setIsLoading3(true);

      try {
        const payload1 = {
          rice_inline: block_data2,
          rice_inline_value: inline_value_rice,
          wheat_inline: data,
          wheat_inline_value: inline_value_wheat,
        };

        const response2 = await fetch(ProjectIp + "/Daily_Planner_Check", {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
          },
          body: JSON.stringify(payload1),
        });

        const responseData1 = await response2.json(); // Parse response JSON

        if (responseData1.status === "NO") {
          alert("Distance is not within range. Please check again.");
        }
      } catch (error) {
        console.error("Error sending inputs:", error);
      } finally {
        // Reset loading state
        setIsLoading3(false);
      }
    }

    // Add progress message
    setProgress((prev) => [
      ...prev,
      "New Inline details have been added for wheat",
    ]);
  };

  const addConstraint3 = async (e) => {
    e.preventDefault();
    console.log(subOption3);

    if (selectedOption3 && subOption3) {
      // Check if the railhead for the selected state is already added
      if (!addedRailheads.has(subOption3)) {
        setBlockdata3((data) => [
          ...data,
          {
            origin_state: selectedOption3,
            origin_railhead: subOption3,
            id: Date.now(),
          },
        ]);

        setnumber_check1(number_check1 + 1);

        // Add the railhead to the set to indicate it's added for this state
        setAddedRailheads(new Set(addedRailheads).add(subOption3));
      } else {
        alert(
          `Railhead ${subOption3} is already added for ${selectedOption3}.`
        );
      }

      setSubOptions3([]);
    }

    setSubOption3("");

    const response = await fetch("/data/Updated_railhead_list.xlsx");
    const arrayBuffer = await response.arrayBuffer();
    const data = new Uint8Array(arrayBuffer);

    const workbook = XLSX.read(data, { type: "array" });

    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    const dropdownOptions = [
      {
        value: "",
        label: "Please select Railhead",
      },
    ];

    for (let i = 0; i < jsonData.length; i++) {
      if (
        jsonData[i][1] &&
        jsonData[i][1].trim().toLowerCase() ===
          selectedOption3.trim().toLowerCase()
      ) {
        dropdownOptions.push({ value: jsonData[i][0], label: jsonData[i][0] });
      }
    }

    dropdownOptions.sort((a, b) => a.label.localeCompare(b.label));

    setSubOptions3(dropdownOptions);
  };

  const addConstraintWheat3 = async (e) => {
    e.preventDefault();
    if (selectedOptionWheat3 && subOptionWheat3) {
      // Check if the railhead for the selected state is already added

      setBlockdataWheat3((data) => [
        ...data,
        {
          origin_state: selectedOptionWheat3,
          origin_railhead: subOptionWheat3,
          id: Date.now(),
        },
      ]);

      setSupplyWeatCount(supplyWeatCount + 1);

      // Add the railhead to the set to indicate it's added for this state
      setAddedRailheads(new Set(addedRailheads).add(subOptionWheat3));

      setSubOptionsWheat3([]);
    }

    setSubOptionWheat3("");

    const response = await fetch("/data/Updated_railhead_list.xlsx");
    const arrayBuffer = await response.arrayBuffer();
    const data = new Uint8Array(arrayBuffer);

    const workbook = XLSX.read(data, { type: "array" });

    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    const dropdownOptions = [
      {
        value: "",
        label: "Please select Railhead",
      },
    ];

    for (let i = 0; i < jsonData.length; i++) {
      if (
        jsonData[i][1] &&
        jsonData[i][1].trim().toLowerCase() ===
          selectedOptionWheat3.trim().toLowerCase()
      ) {
        dropdownOptions.push({ value: jsonData[i][0], label: jsonData[i][0] });
      }
    }

    dropdownOptions.sort((a, b) => a.label.localeCompare(b.label));

    setSubOptionsWheat3(dropdownOptions);
  };

  const addConstraint4 = async (e) => {
    e.preventDefault();

    if (selectedOption4 && subOption4) {
      // Check if the railhead for the selected state is already added
      if (!addedRailheads.has(subOption4)) {
        setRiceDestination((data) => [
          ...data,
          {
            origin_state: selectedOption4,
            origin_railhead: subOption4,
            id: Date.now(),
          },
        ]);

        setnumber_check2(number_check2 + 1);

        // Add the railhead to the set to indicate it's added for this state
        setAddedRailheads(new Set(addedRailheads).add(subOption4));
      } else {
        alert(
          `Railhead ${subOption4} is already added for ${selectedOption4}.`
        );
      }

      setSubOptions4([]);
    }

    setSubOption4("");

    const response = await fetch("/data/Updated_railhead_list.xlsx");
    const arrayBuffer = await response.arrayBuffer();
    const data = new Uint8Array(arrayBuffer);

    const workbook = XLSX.read(data, { type: "array" });

    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    const dropdownOptions = [
      {
        value: "",
        label: "Please select Railhead",
      },
    ];

    for (let i = 0; i < jsonData.length; i++) {
      if (
        jsonData[i][1] &&
        jsonData[i][1].trim().toLowerCase() ===
          selectedOption4.trim().toLowerCase()
      ) {
        dropdownOptions.push({ value: jsonData[i][0], label: jsonData[i][0] });
      }
    }

    dropdownOptions.sort((a, b) => a.label.localeCompare(b.label));

    setSubOptions4(dropdownOptions);
  };

  const addConstraintWheat4 = async (e) => {
    e.preventDefault();
    if (selectedOptionWheat4 && subOptionWheat4) {
      // Check if the railhead for the selected state is already added
      if (!addedRailheads.has(subOptionWheat4)) {
        setWheatDestination((data) => [
          ...data,
          {
            origin_state: selectedOptionWheat4,
            origin_railhead: subOptionWheat4,
            id: Date.now(),
          },
        ]);

        setDestinationWheatCount(destinationWheatCount + 1);

        // Add the railhead to the set to indicate it's added for this state
        setAddedRailheads(new Set(addedRailheads).add(subOptionWheat4));
      } else {
        // Alert message if the railhead is already added for this state
        alert(
          `Railhead ${subOptionWheat4} is already added for ${selectedOptionWheat4}.`
        );
      }

      setSubOptionsWheat4([]);
    }

    setSubOptionWheat4("");

    const response = await fetch("/data/Updated_railhead_list.xlsx");
    const arrayBuffer = await response.arrayBuffer();
    const data = new Uint8Array(arrayBuffer);

    const workbook = XLSX.read(data, { type: "array" });

    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    const dropdownOptions = [
      {
        value: "",
        label: "Please select Railhead",
      },
    ];

    for (let i = 0; i < jsonData.length; i++) {
      if (
        jsonData[i][1] &&
        jsonData[i][1].trim().toLowerCase() ===
          selectedOptionWheat4.trim().toLowerCase()
      ) {
        dropdownOptions.push({ value: jsonData[i][0], label: jsonData[i][0] });
      }
    }

    dropdownOptions.sort((a, b) => a.label.localeCompare(b.label));

    setSubOptionsWheat4(dropdownOptions);
  };

  const addConstraint_fixed = (e) => {
    e.preventDefault();
    if (
      selectedOption_fixed &&
      subOption1_fixed &&
      selectedOption2_fixed &&
      subOption2_fixed &&
      commodity_fixed
      // && value_fixed
    ) {
      // Check if origin and destination railheads are the same
      // if (subOption1_fixed === subOption2_fixed) {
      //   alert("Origin and destination railheads cannot be the same.");
      //   return; // Do not proceed further
      // }

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

      // Reset options and suboptions
      setSelectedOption_fixed("default");
      setSelectedOption2_fixed("default");
      setSubOptions_fixed([]);
      setSubOptions2_fixed([]);

      setProgress((prev) => [
        ...prev,
        `Route from ${subOption1_fixed} to ${subOption2_fixed} has been fixed for ${commodity_fixed}`,
      ]);
    }
  };

  const viewGrid = () => {
    setShowMessage(true);
    const riceData = JSON.parse(Total_result?.rice ?? 0);
    console.log("ricedata", riceData);
    const wheatData = JSON.parse(Total_result?.wheat ?? 0);
    console.log("wheatData", wheatData);
    setRiceData(riceData);
    setWheatData(wheatData);
  };

  const exportToExcel1 = () => {
    if (Total_result == null) {
      // Commented out the alert statement
      window.alert("Fetching Result, Please Wait");
      fetchReservationId_Total_result();
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

      const currentDate = new Date();
      const year = currentDate.getFullYear();
      const month = String(currentDate.getMonth() + 1).padStart(2, "0"); // Months are zero-indexed, so we add 1
      const day = String(currentDate.getDate()).padStart(2, "0");
      const hours = String(currentDate.getHours()).padStart(2, "0");
      const minutes = String(currentDate.getMinutes()).padStart(2, "0");
      const seconds = String(currentDate.getSeconds()).padStart(2, "0");
      const dateAndTime = `${year}/${month}/${day}T${hours}/${minutes}/${seconds}`;
      const filenameWithDateTime = `Daily_Movement_Scenario1_${dateAndTime}.xlsx`;
      saveAs(excelBlob, filenameWithDateTime);
    }
  };

  const exportToExcel2 = () => {
    if (Relevant_result == null) {
      // Commented out the alert statement
      window.alert("Fetching Result, Please Wait");
      fetchReservationId_Revelant_result();
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
      // Commented out the alert statement
      // window.alert("Result Downloaded");
    }
  };

  const buttonStyle = {
    border: updateExcel ? "4px solid rgba(235, 171, 68)" : "2px solid black",
    padding: "5px",
  };

  const buttonStyle2 = {
    border: updateExcel2 ? "4px solid rgba(235, 171, 68)" : "2px solid black",
    padding: "5px",
  };

  return (
    <div className="page-container" style={{ backgroundColor: "#ebab44b0" }}>
      <Sidenav />
      <div
        className="page-content"
        style={{
          display: "flex",
          backgroundImage: "url('static/img/bg8.jpg')",
        }}
      >
        <div>
          <ul
            className="x-navigation x-navigation-horizontal x-navigation-panel"
            style={{ backgroundColor: "rgba(235, 171, 68, 0.69)" }}
          >
            <li className="xn-icon-button">
              <a href="javascript:void(0)" className="x-navigation-minimize">
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
                  {/* <div style={{ fontSize: "20px", fontWeight: "700" }}>
                    <i className="fa fa-file-excel-o" aria-hidden="true"></i>{" "}
                    Template
                  </div> */}
                  {/* <form
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
                </form> */}
                </div>
                <br />
                <br />
                <div style={{ display: "flex", marginLeft: "245px" }}>
                  {/* {fileSelected && ( */}
                  {/* <div style={{ marginTop: "-20px" }}>
                    <button style={buttonStyle} onClick={() => update_excel()}>
                      Template 1
                    </button>
                  </div> */}
                  {/* )} */}
                  {/* {updateExcel && ( */}
                  {/* <div style={{ marginLeft: "150px", marginTop: "-20px" }}>
                    <button
                      style={buttonStyle2}
                      onClick={() => update_excel2()}
                    >
                      Template for Scenario 2
                    </button>
                  </div> */}
                  {/* )} */}
                </div>
                {/* {(updateExcel || updateExcel2) && (
                  <div style={{ marginLeft: "480px" }}>
                    <br />
                    <button
                      style={{ padding: "5px" }}
                      onClick={() => save_excel()}
                    >
                      Save changes
                    </button>
                  </div>
                )} */}
                {activeSheetName &&
                  (updateExcel || updateExcel2) &&
                  excelData[activeSheetName] && (
                    <div style={{ marginLeft: "20%" }}>
                      <select
                        onChange={(e) => setActiveSheetName(e.target.value)}
                        value={activeSheetName}
                        style={{
                          margin: "10px",
                          padding: "3px",
                          marginLeft: "20%",
                        }}
                      >
                        {Object.keys(excelData).map((sheetName) => (
                          <option key={sheetName} value={sheetName}>
                            {sheetName}
                          </option>
                        ))}
                      </select>
                      <table>
                        <tbody>
                          {excelData[activeSheetName].map((row, rowIndex) => (
                            <tr key={rowIndex}>
                              {row.map((cellValue, columnIndex) => (
                                <td key={columnIndex}>
                                  <input
                                    type="text"
                                    value={cellValue}
                                    onChange={(e) =>
                                      handleCellChange(
                                        activeSheetName,
                                        rowIndex,
                                        columnIndex,
                                        e.target.value
                                      )
                                    }
                                  />
                                </td>
                              ))}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  )}
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
                        onChange={(e) => {
                          set_TEFD(e.target.value);
                          setProgress((prev) => [
                            ...prev,
                            `You have selected the matrix system as ${e.target.value}`,
                          ]);
                        }}
                        style={{ marginLeft: "547px" }}
                      >
                        <option>Select Matrix System</option>
                        <option value="NON-TEFD">Non-TEFD</option>
                        <option value="TEFD">TEFD</option>
                        <option value="Non-TEFD+TC">Non-TEFD + TC</option>
                        <option value="TEFD+TC">TEFD + TC</option>
                      </select>
                    </label>
                    {/* <label>
                      <strong
                        style={{
                          fontSize: "20px",
                          marginLeft: "15px",
                          color: "#9d0921",
                        }}
                      >
                        Select Scenario
                      </strong>
                      <select
                        value={Scenerio}
                        onChange={(e) => {
                          set_Scenerio(e.target.value);
                          document.getElementById("console_").style.display =
                            "block";
                          document.getElementById("console_").innerHTML +=
                            "You have selected the scenario as " +
                            e.target.value +
                            "<br/><br/>";
                        }}
                        style={{ marginLeft: "600px" }}
                      >
                        <option value="">Select Scenario</option>
                        <option value="Scenerio 1">Scenario 1</option>
                        <option value="Scenerio 2">Scenario 2</option>
                      </select>
                    </label> */}
                    <br />
                    <br />
                    <div>
                      <p style={{ margin: 0, padding: 0 }}>
                        <strong
                          style={{
                            color: "#9d0921",
                            fontSize: "20px",
                            marginLeft: "15px",
                          }}
                        >
                          For Rice:
                        </strong>
                      </p>
                      <br />
                      <div
                        style={{
                          display: "flex",
                          marginLeft: "20px",
                          alignItems: "center",
                        }}
                      >
                        <div>
                          <strong style={{ fontSize: "16px", padding: "5px" }}>
                            Select Origin State:
                          </strong>
                          <select
                            style={{
                              width: "200px",
                              padding: "5px",
                              marginRight: 25,
                            }}
                            onChange={handleDropdownChange3}
                            value={selectedOption3}
                          >
                            <option value="default">Select Origin State</option>
                            <option value="Andhra Pradesh">
                              Andhra Pradesh
                            </option>
                            <option value="Bihar">Bihar</option>
                            <option value="Chattisgarh">Chattisgarh</option>
                            <option value="Goa">Goa</option>
                            <option value="Gujarat">Gujarat</option>
                            <option value="Haryana">Haryana</option>
                            <option value="Jammu & Kashmir">
                              Jammu & Kashmir
                            </option>
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
                            style={{
                              width: "200px",
                              padding: "5px",
                            }}
                            onChange={handleSubDropdownChange3}
                            value={subOption3}
                          >
                            {subOptions3.map((option) => (
                              <option key={option.value} value={option.value}>
                                {option.label}
                              </option>
                            ))}
                          </select>
                        </div>
                        <div onClick={addConstraint3}>
                          <button
                            style={{
                              textAlign: "center",
                              backgroundColor: "orange",
                              width: 70,
                              height: 40,
                              marginLeft: 346,
                            }}
                            disabled={
                              subOption3 === "" && selectedOption3 === ""
                            }
                          >
                            Add
                          </button>
                        </div>
                      </div>

                      <br />

                      {block_data3.length !== 0 && (
                        <div>
                          <table>
                            <thead>
                              <tr style={{ margin: "auto" }}>
                                <th style={{ padding: "10px", width: "396px" }}>
                                  Origin State
                                </th>
                                <th style={{ padding: "10px", width: "396px" }}>
                                  Origin Railhead
                                </th>
                                <th style={{ padding: "10px", width: "396px" }}>
                                  Delete
                                </th>
                              </tr>
                            </thead>
                            <tbody>
                              {block_data3.map((item) => (
                                <tr key={item.id}>
                                  <td>{item.origin_state}</td>
                                  <td>{item.origin_railhead}</td>

                                  <td>
                                    <span
                                      style={{
                                        cursor: "pointer",
                                        color: "#ff0000",
                                        fontSize: "1.2rem",
                                      }}
                                      onClick={() =>
                                        handleDeleteRow_Rice_s(item.id)
                                      }
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
                      <div
                        style={{
                          display: "flex",
                          marginLeft: "20px",
                          alignItems: "center",
                        }}
                      >
                        <div>
                          <strong style={{ fontSize: "16px", padding: "5px" }}>
                            Select Destination State:
                          </strong>
                          <select
                            style={{
                              width: "200px",
                              padding: "5px",
                              marginRight: 25,
                            }}
                            id="destination"
                            onChange={handleDropdownChange4}
                            value={selectedOption4}
                          >
                            <option value="default">
                              Select Destination State
                            </option>
                            <option value="Andhra Pradesh">
                              Andhra Pradesh
                            </option>
                            <option value="Bihar">Bihar</option>
                            <option value="Chattisgarh">Chattisgarh</option>
                            <option value="Goa">Goa</option>
                            <option value="Gujarat">Gujarat</option>
                            <option value="Haryana">Haryana</option>
                            <option value="Jammu & Kashmir">
                              Jammu & Kashmir
                            </option>
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
                            Select Destination Railhead
                          </strong>
                          <select
                            style={{
                              width: "200px",
                              padding: "5px",
                            }}
                            onChange={handleSubDropdownChange4}
                            value={subOption4}
                          >
                            {subOptions4.map((option) => (
                              <option key={option.value} value={option.value}>
                                {option.label}
                              </option>
                            ))}
                          </select>
                        </div>
                        <div onClick={addConstraint4}>
                          <button
                            style={{
                              textAlign: "center",
                              backgroundColor: "orange",
                              width: 70,
                              height: 40,
                              alignItems: "center",
                              marginLeft: 267,
                            }}
                            disabled={
                              subOption4 === "" && selectedOption4 === ""
                            }
                          >
                            Add
                          </button>
                        </div>
                      </div>
                      <br />
                      <div>
                        {rice_destination.length !== 0 && (
                          <div>
                            <table>
                              <thead>
                                <tr style={{ margin: "auto" }}>
                                  <th
                                    style={{ padding: "10px", width: "396px" }}
                                  >
                                    Destination State
                                  </th>
                                  <th
                                    style={{ padding: "10px", width: "396px" }}
                                  >
                                    Destination Railhead
                                  </th>
                                  <th
                                    style={{ padding: "10px", width: "396px" }}
                                  >
                                    Delete
                                  </th>
                                </tr>
                              </thead>
                              <tbody>
                                {rice_destination.map((item, index) => (
                                  <tr key={index}>
                                    <td>{item.origin_state}</td>
                                    <td>{item.origin_railhead}</td>
                                    <td>
                                      <span
                                        style={{
                                          cursor: "pointer",
                                          color: "#ff0000",
                                          fontSize: "1.2rem",
                                        }}
                                        onClick={() =>
                                          handleDeleteRow_Rice__dest(index)
                                        }
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
                      <br />
                      <br />
                      <div>
                        <div style={{ marginLeft: "15px" }}>
                          <strong style={{ fontSize: "16px", padding: "5px" }}>
                            Enter Inline Value
                          </strong>
                          <input
                            type="number"
                            value={inline_value_rice}
                            onChange={(e) => setInlineValueRice(e.target.value)}
                            style={{
                              marginLeft: "40px",
                              width: "200px",
                              padding: "5px",
                            }}
                          />
                        </div>
                        <br />
                        <div
                          style={{
                            display: "flex",
                            marginLeft: "20px",
                            width: "1170px",
                            alignItems: "center",
                          }}
                        >
                          <div>
                            <strong
                              style={{ fontSize: "16px", padding: "5px" }}
                            >
                              Select Inline State
                            </strong>
                            <select
                              style={{ width: "200px", padding: "5px" }}
                              onChange={handleDropdownChange5}
                              value={selectedOption5}
                            >
                              <option value="default">
                                Select Inline State
                              </option>
                              <option value="Andhra Pradesh">
                                Andhra Pradesh
                              </option>
                              <option value="Bihar">Bihar</option>
                              <option value="Chattisgarh">Chattisgarh</option>
                              <option value="Goa">Goa</option>
                              <option value="Gujarat">Gujarat</option>
                              {/* <option value="Haryana">Haryana</option> */}
                              <option value="Jammu & Kashmir">
                                Jammu & Kashmir
                              </option>
                              <option value="Jharkhand">Jharkhand</option>
                              <option value="Karnataka">Karnataka</option>
                              <option value="Kerala">Kerala</option>
                              {/* <option value="MP">Madhya Pradesh</option> */}
                              <option value="Maharashtra">Maharashtra</option>
                              <option value="NE">North East</option>
                              <option value="Odisha">Odisha</option>
                              {/* <option value="Punjab">Punjab</option> */}
                              <option value="Rajasthan">Rajasthan</option>
                              <option value="Tamil Nadu">Tamil Nadu</option>
                              <option value="Telangana">Telangana</option>
                              <option value="UP">Uttar Pradesh</option>
                              <option value="Uttarakhand">Uttarakhand</option>
                              <option value="West Bengal">West Bengal</option>
                            </select>
                          </div>

                          <div>
                            <strong
                              style={{ fontSize: "16px", padding: "5px" }}
                            >
                              Select Inline Railhead
                            </strong>
                            <select
                              style={{ width: "200px", padding: "5px" }}
                              onChange={handleSubDropdownChange5}
                              value={
                                subOption5 === "default"
                                  ? "default"
                                  : subOption5
                              }
                            >
                              {subOptions5.map((option) => (
                                <option key={option.value} value={option.value}>
                                  {option.label}
                                </option>
                              ))}
                            </select>
                          </div>

                          <div>
                            <strong
                              style={{ fontSize: "16px", padding: "5px" }}
                            >
                              Select Inline State
                            </strong>
                            <select
                              style={{ width: "200px", padding: "5px" }}
                              onChange={handleDropdownChange6}
                              value={selectedOption6}
                            >
                              <option value="default">
                                Select Inline State
                              </option>
                              <option value="Andhra Pradesh">
                                Andhra Pradesh
                              </option>
                              <option value="Bihar">Bihar</option>
                              <option value="Chattisgarh">Chattisgarh</option>
                              <option value="Goa">Goa</option>
                              <option value="Gujarat">Gujarat</option>
                              <option value="Jammu & Kashmir">
                                Jammu & Kashmir
                              </option>
                              <option value="Jharkhand">Jharkhand</option>
                              <option value="Karnataka">Karnataka</option>
                              <option value="Kerala">Kerala</option>
                              {/* <option value="MP">Madhya Pradesh</option> */}
                              <option value="Maharashtra">Maharashtra</option>
                              <option value="NE">North East</option>
                              <option value="Odisha">Odisha</option>
                              {/* <option value="Punjab">Punjab</option> */}
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
                            <strong
                              style={{ fontSize: "16px", padding: "5px" }}
                            >
                              Select Inline Railhead
                            </strong>
                            <select
                              style={{ width: "200px", padding: "5px" }}
                              onChange={handleSubDropdownChange6}
                              value={subOption6}
                            >
                              {subOptions6.map((option) => (
                                <option key={option.value} value={option.value}>
                                  {option.label}
                                </option>
                              ))}
                            </select>
                          </div>
                          {/* </label> */}
                          <div onClick={addConstraint2}>
                            <button
                              style={{
                                textAlign: "center",
                                backgroundColor: "orange",
                                width: 70,
                                height: 40,
                                alignItems: "center",
                              }}
                            >
                              Add
                            </button>
                          </div>
                        </div>
                      </div>
                      <br />

                      {block_data2.length !== 0 && (
                        <div>
                          {/* <div
                      style={{
                        fontSize: "20px",
                        fontWeight: "700",
                        textAlign: "center",
                      }}
                    >
                      Route Block
                    </div> */}
                          <table>
                            <thead>
                              <tr style={{ margin: "auto" }}>
                                <th style={{ padding: "10px", width: "238px" }}>
                                  Inline State
                                </th>
                                <th style={{ padding: "10px", width: "238px" }}>
                                  Inline Railhead
                                </th>
                                <th style={{ padding: "10px", width: "238px" }}>
                                  Inline State
                                </th>
                                <th style={{ padding: "10px", width: "238px" }}>
                                  Inline Railhead
                                </th>
                                <th style={{ padding: "10px", width: "236px" }}>
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
                              {block_data2.map((item) => (
                                <tr key={item.id}>
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
                                      onClick={() =>
                                        handleDeleteRow_inline(item.id)
                                      }
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
                    <br />
                    <div>
                      <p style={{ margin: 0, padding: 0 }}>
                        <strong
                          style={{
                            color: "#9d0921",
                            fontSize: "20px",
                            marginLeft: "15px",
                          }}
                        >
                          For Wheat:
                        </strong>
                      </p>
                      <br />
                      <div
                        style={{
                          display: "flex",
                          marginLeft: "20px",
                          alignItems: "center",
                        }}
                      >
                        <div>
                          <strong style={{ fontSize: "16px", padding: "5px" }}>
                            Select Origin State:
                          </strong>
                          <select
                            style={{
                              width: "200px",
                              padding: "5px",
                              marginRight: 25,
                            }}
                            onChange={handleDropdownChangeWheat3}
                            value={selectedOptionWheat3}
                          >
                            <option value="default">Select Origin State</option>
                            <option value="Andhra Pradesh">
                              Andhra Pradesh
                            </option>
                            <option value="Bihar">Bihar</option>
                            <option value="Chattisgarh">Chattisgarh</option>
                            <option value="Goa">Goa</option>
                            <option value="Gujarat">Gujarat</option>
                            <option value="Haryana">Haryana</option>
                            <option value="Jammu & Kashmir">
                              Jammu & Kashmir
                            </option>
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
                            style={{
                              width: "200px",
                              padding: "5px",
                            }}
                            onChange={handleSubDropdownChangeWheat3}
                            value={subOptionWheat3}
                          >
                            {subOptionsWheat3.map((option) => (
                              <option key={option.value} value={option.value}>
                                {option.label}
                              </option>
                            ))}
                          </select>
                        </div>
                        <div onClick={addConstraintWheat3}>
                          <button
                            style={{
                              textAlign: "center",
                              backgroundColor: "orange",
                              width: 70,
                              height: 40,
                              marginLeft: 346,
                            }}
                            disabled={
                              subOptionWheat3 === "" &&
                              selectedOptionWheat3 === ""
                            }
                          >
                            Add
                          </button>
                        </div>
                      </div>

                      <br />

                      {block_dataWheat3.length !== 0 && (
                        <div>
                          <table>
                            <thead>
                              <tr style={{ margin: "auto" }}>
                                <th style={{ padding: "10px", width: "396px" }}>
                                  Origin State
                                </th>
                                <th style={{ padding: "10px", width: "396px" }}>
                                  Origin Railhead
                                </th>
                                <th style={{ padding: "10px", width: "396px" }}>
                                  Delete
                                </th>
                              </tr>
                            </thead>
                            <tbody>
                              {block_dataWheat3.map((item) => (
                                <tr key={item.id}>
                                  <td>{item.origin_state}</td>
                                  <td>{item.origin_railhead}</td>

                                  <td>
                                    <span
                                      style={{
                                        cursor: "pointer",
                                        color: "#ff0000",
                                        fontSize: "1.2rem",
                                      }}
                                      onClick={() =>
                                        handleDeleteRow_Wheat_s(item.id)
                                      }
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

                      <div
                        style={{
                          display: "flex",
                          marginLeft: "20px",
                          alignItems: "center",
                        }}
                      >
                        <div>
                          <strong style={{ fontSize: "16px", padding: "5px" }}>
                            Select Destination State:
                          </strong>
                          <select
                            style={{ width: "200px", padding: "5px" }}
                            id="destination"
                            onChange={handleDropdownChangeWheat4}
                            value={selectedOptionWheat4}
                          >
                            <option value="default">
                              Select Destination State
                            </option>
                            <option value="Andhra Pradesh">
                              Andhra Pradesh
                            </option>
                            <option value="Bihar">Bihar</option>
                            <option value="Chattisgarh">Chattisgarh</option>
                            <option value="Goa">Goa</option>
                            <option value="Gujarat">Gujarat</option>
                            <option value="Haryana">Haryana</option>
                            <option value="Jammu & Kashmir">
                              Jammu & Kashmir
                            </option>
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
                          <strong
                            style={{
                              fontSize: "16px",
                              padding: "5px",
                              marginLeft: 15,
                            }}
                          >
                            Select Destination Railhead
                          </strong>
                          <select
                            style={{ width: "200px", padding: "5px" }}
                            onChange={handleSubDropdownChangeWheat4}
                            value={subOptionWheat4}
                          >
                            {subOptionsWheat4.map((option) => (
                              <option key={option.value} value={option.value}>
                                {option.label}
                              </option>
                            ))}
                          </select>
                        </div>
                        <div onClick={addConstraintWheat4}>
                          <button
                            style={{
                              textAlign: "center",
                              backgroundColor: "orange",
                              width: 70,
                              height: 40,
                              marginLeft: 277,
                            }}
                            disabled={
                              subOptionWheat4 === "" &&
                              selectedOptionWheat4 === ""
                            }
                          >
                            Add
                          </button>
                        </div>
                      </div>
                      <br />
                      <div>
                        {wheat_destination.length !== 0 && (
                          <div>
                            <table>
                              <thead>
                                <tr style={{ margin: "auto" }}>
                                  <th
                                    style={{ padding: "10px", width: "396px" }}
                                  >
                                    Destination State
                                  </th>
                                  <th
                                    style={{ padding: "10px", width: "396px" }}
                                  >
                                    Destination Railhead
                                  </th>
                                  <th
                                    style={{ padding: "10px", width: "396px" }}
                                  >
                                    Delete
                                  </th>
                                </tr>
                              </thead>
                              <tbody>
                                {wheat_destination.map((item, index) => (
                                  <tr key={index}>
                                    <td>{item.origin_state}</td>
                                    <td>{item.origin_railhead}</td>

                                    <td>
                                      <span
                                        style={{
                                          cursor: "pointer",
                                          color: "#ff0000",
                                          fontSize: "1.2rem",
                                        }}
                                        onClick={() =>
                                          handleDeleteRow_Wheat__dest(index)
                                        }
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
                      <br />
                      <br />
                      <div>
                        <div style={{ marginLeft: "15px" }}>
                          <strong style={{ fontSize: "16px", padding: "5px" }}>
                            Enter Inline Value
                          </strong>
                          <input
                            type="number"
                            value={inline_value_wheat}
                            onChange={(e) =>
                              setInlineValueWheat(e.target.value)
                            }
                            style={{
                              marginLeft: "40px",
                              width: "200px",
                              padding: "5px",
                            }}
                          />
                        </div>
                        <br />
                        <div
                          style={{
                            display: "flex",
                            marginLeft: "20px",
                            width: "1170px",
                            alignItems: "center",
                          }}
                        >
                          {/* <label htmlFor="origin_state"> */}
                          <div>
                            <strong
                              style={{ fontSize: "16px", padding: "5px" }}
                            >
                              Select Inline State
                            </strong>
                            <select
                              style={{ width: "200px", padding: "5px" }}
                              onChange={handleDropdownChangeWheat5}
                              value={selectedOptionWheat5}
                            >
                              <option value="default">
                                Select Inline State
                              </option>
                              <option value="Andhra Pradesh">
                                Andhra Pradesh
                              </option>
                              <option value="Bihar">Bihar</option>
                              <option value="Chattisgarh">Chattisgarh</option>
                              <option value="Goa">Goa</option>
                              <option value="Gujarat">Gujarat</option>
                              {/* <option value="Haryana">Haryana</option> */}
                              <option value="Jammu & Kashmir">
                                Jammu & Kashmir
                              </option>
                              <option value="Jharkhand">Jharkhand</option>
                              <option value="Karnataka">Karnataka</option>
                              <option value="Kerala">Kerala</option>
                              {/* <option value="MP">Madhya Pradesh</option> */}
                              <option value="Maharashtra">Maharashtra</option>
                              <option value="NE">North East</option>
                              <option value="Odisha">Odisha</option>
                              {/* <option value="Punjab">Punjab</option> */}
                              <option value="Rajasthan">Rajasthan</option>
                              <option value="Tamil Nadu">Tamil Nadu</option>
                              <option value="Telangana">Telangana</option>
                              <option value="UP">Uttar Pradesh</option>
                              <option value="Uttarakhand">Uttarakhand</option>
                              <option value="West Bengal">West Bengal</option>
                            </select>
                          </div>

                          <div>
                            <strong
                              style={{ fontSize: "16px", padding: "5px" }}
                            >
                              Select Inline Railhead
                            </strong>
                            <select
                              style={{ width: "200px", padding: "5px" }}
                              onChange={handleSubDropdownChangeWheat5}
                              value={subOptionWheat5}
                            >
                              {subOptionsWheat5.map((option) => (
                                <option key={option.value} value={option.value}>
                                  {option.label}
                                </option>
                              ))}
                            </select>
                          </div>
                          {/* </label> */}
                          <div>
                            {/* <label htmlFor="deficit_state"> */}
                            <strong
                              style={{ fontSize: "16px", padding: "5px" }}
                            >
                              Select Inline State
                            </strong>
                            <select
                              style={{ width: "200px", padding: "5px" }}
                              onChange={handleDropdownChangeWheat6}
                              value={selectedOptionWheat6}
                            >
                              <option value="default">
                                Select Inline State
                              </option>
                              <option value="Andhra Pradesh">
                                Andhra Pradesh
                              </option>
                              <option value="Bihar">Bihar</option>
                              <option value="Chattisgarh">Chattisgarh</option>
                              <option value="Goa">Goa</option>
                              <option value="Gujarat">Gujarat</option>
                              {/* <option value="Haryana">Haryana</option> */}
                              <option value="Jammu & Kashmir">
                                Jammu & Kashmir
                              </option>
                              <option value="Jharkhand">Jharkhand</option>
                              <option value="Karnataka">Karnataka</option>
                              <option value="Kerala">Kerala</option>
                              {/* <option value="MP">Madhya Pradesh</option> */}
                              <option value="Maharashtra">Maharashtra</option>
                              <option value="NE">North East</option>
                              <option value="Odisha">Odisha</option>
                              {/* <option value="Punjab">Punjab</option> */}
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
                            <strong
                              style={{ fontSize: "16px", padding: "5px" }}
                            >
                              Select Inline Railhead
                            </strong>
                            <select
                              style={{ width: "200px", padding: "5px" }}
                              onChange={handleSubDropdownChangeWheat6}
                              value={subOptionWheat6}
                            >
                              {subOptionsWheat6.map((option) => (
                                <option key={option.value} value={option.value}>
                                  {option.label}
                                </option>
                              ))}
                            </select>
                          </div>
                          {/* </label> */}
                          <div onClick={addConstraintWheat2}>
                            <button
                              style={{
                                textAlign: "center",
                                backgroundColor: "orange",
                                width: 70,
                                height: 40,
                                alignItems: "center",
                              }}
                            >
                              Add
                            </button>
                          </div>
                        </div>
                      </div>
                      <br />

                      {block_dataWheat2.length !== 0 && (
                        <div>
                          {/* <div
                      style={{
                        fontSize: "20px",
                        fontWeight: "700",
                        textAlign: "center",
                      }}
                    >
                      Route Block
                    </div> */}
                          <table>
                            <thead>
                              <tr style={{ margin: "auto" }}>
                                <th style={{ padding: "10px", width: "238px" }}>
                                  Inline State
                                </th>
                                <th style={{ padding: "10px", width: "238px" }}>
                                  Inline Railhead
                                </th>
                                <th style={{ padding: "10px", width: "238px" }}>
                                  Inline State
                                </th>
                                <th style={{ padding: "10px", width: "238px" }}>
                                  Inline Railhead
                                </th>
                                <th style={{ padding: "10px", width: "236px" }}>
                                  Delete
                                </th>
                              </tr>
                              {/* <tr  style={{ padding: "10px", width: "100%" , textAlign:'center'}}>
                      <div style={{textAlign:'center', width:'100%'}}>Routes Block</div></tr> */}
                            </thead>
                            <tbody>
                              {block_dataWheat2.map((item) => (
                                <tr key={item.id}>
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
                                      onClick={() =>
                                        handleDeleteRow_inlineWheat(item.id)
                                      }
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
                    <div
                      style={{
                        display: "flex",
                        marginLeft: "20px",
                        width: 1170,
                      }}
                    >
                      {/* <label htmlFor="origin_state"> */}
                      <div>
                        <strong style={{ fontSize: "16px", padding: "5px" }}>
                          Select Origin State
                        </strong>
                        <select
                          style={{ width: "200px", padding: "5px" }}
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
                          <option value="Jammu & Kashmir">
                            Jammu & Kashmir
                          </option>
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
                      <div>
                        <strong style={{ fontSize: "16px", padding: "5px" }}>
                          Select Destination State
                        </strong>
                        <select
                          style={{ width: "200px", padding: "5px" }}
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
                          <option value="Jammu & Kashmir">
                            Jammu & Kashmir
                          </option>
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
                          Select Destination Railhead
                        </strong>
                        <select
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
                      <div onClick={addConstraint}>
                        <button
                          style={{
                            textAlign: "center",
                            backgroundColor: "orange",
                            width: 70,
                            height: 40,
                            alignItems: "center",
                          }}
                        >
                          Add
                        </button>
                      </div>
                    </div>
                    <br />
                    {!solutionSolved && block_data.length !== 0 && (
                      <div>
                        <table>
                          <thead>
                            <tr style={{ margin: "auto" }}>
                              <th style={{ padding: "10px", width: "238px" }}>
                                Origin State
                              </th>
                              <th style={{ padding: "10px", width: "238px" }}>
                                Origin Railhead
                              </th>
                              <th style={{ padding: "10px", width: "238px" }}>
                                Destination State
                              </th>
                              <th style={{ padding: "10px", width: "238px" }}>
                                Destination Railhead
                              </th>
                              <th style={{ padding: "10px", width: "236px" }}>
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
                              <tr key={item.id}>
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
                    <div
                      style={{
                        display: "flex",
                        marginLeft: "20px",
                        width: 1030,
                      }}
                    >
                      {/* <label htmlFor="origin_state"> */}
                      <div>
                        <strong style={{ fontSize: "16px", padding: "5px" }}>
                          Select Origin State
                        </strong>
                        <select
                          style={{ width: "200px", padding: "5px" }}
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
                          <option value="Jammu & Kashmir">
                            Jammu & Kashmir
                          </option>
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
                          <option value="Jammu & Kashmir">
                            Jammu & Kashmir
                          </option>
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
                    </div>
                    <div
                      style={{
                        display: "flex",
                        marginLeft: "20px",
                        marginTop: "10px",
                      }}
                    >
                      <strong style={{ fontSize: "16px", padding: "5px" }}>
                        Select Commodity
                      </strong>
                      <select
                        value={commodity_fixed}
                        onChange={(e) => setCommodity_fixed(e.target.value)}
                        style={{
                          marginLeft: "40px",
                          width: "200px",
                          padding: "5px",
                        }}
                      >
                        <option value="">Select Commodity</option>
                        <option value="RICE">Rice</option>
                        <option value="WHEAT">Wheat</option>
                      </select>
                      {/* <div style={{ marginLeft: "50px" }}>
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
                      </div> */}

                      <div onClick={addConstraint_fixed}>
                        <button
                          style={{
                            textAlign: "center",
                            backgroundColor: "orange",
                            width: 70,
                            height: 40,
                            marginLeft: 713,
                          }}
                        >
                          Add
                        </button>
                      </div>
                    </div>
                    <br />
                    {!solutionSolved && fixed_data.length !== 0 && (
                      <div>
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
                              {/* <th style={{ padding: "10px", width: "15%" }}>
                                Value
                              </th> */}
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
                              <tr key={item.id}>
                                <td>{item.origin_state}</td>
                                <td>{item.origin_railhead}</td>
                                <td>{item.destination_state}</td>
                                <td>{item.destination_railhead}</td>
                                <td>{item.commodity}</td>
                                {/* <td>{item.value}</td> */}
                                <td>
                                  <span
                                    style={{
                                      cursor: "pointer",
                                      color: "#ff0000",
                                      fontSize: "1.2rem",
                                    }}
                                    onClick={() =>
                                      handleDeleteRow_fixed(item.id)
                                    }
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
                  </form>

                  <div style={{ fontSize: "20px", fontWeight: "700" }}>
                    <i className="fa fa-list-alt" aria-hidden="true"></i>{" "}
                    Optimal Plan
                  </div>
                  <div
                    className="wrap__toggle"
                    style={{
                      textAlign: "center",
                      borderStyle: "solid",
                      borderColor: "#ebab44b0",
                    }}
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
                          onChange={handleSolve}
                        />
                        <span></span>
                      </label>
                    </div>
                  </div>
                  <br />
                  <br />
                  <br />
                  <br />
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

                          <button
                            style={{ color: "white", marginLeft: "15px" }}
                            className="btn btn-danger dropdown-toggle"
                            onClick={viewGrid}
                          >
                            View Railhead Detailed Plan
                          </button>
                          <div>
                            {showMessage && (
                              <div>
                                {riceData !== null && riceData.length > 0 ? (
                                  <div>
                                    <table>
                                      <thead>
                                        <tr style={{ margin: "auto" }}>
                                          <th
                                            style={{
                                              padding: "10px",
                                              width: "350px",
                                            }}
                                          >
                                            Sr. No
                                          </th>
                                          <th
                                            style={{
                                              padding: "10px",
                                              width: "350px",
                                            }}
                                          >
                                            Src RH
                                          </th>
                                          <th
                                            style={{
                                              padding: "10px",
                                              width: "350px",
                                            }}
                                          >
                                            Src state
                                          </th>
                                          <th
                                            style={{
                                              padding: "10px",
                                              width: "350px",
                                            }}
                                          >
                                            Dest RH
                                          </th>
                                          <th
                                            style={{
                                              padding: "10px",
                                              width: "350px",
                                            }}
                                          >
                                            Dest State
                                          </th>
                                          <th
                                            style={{
                                              padding: "10px",
                                              width: "350px",
                                            }}
                                          >
                                            commodity
                                          </th>
                                          {/* <th style={{ padding: "10px", width: "350px" }}>
                                        values
                                      </th> */}
                                        </tr>
                                      </thead>
                                      <tbody>
                                        {riceData.map((item, index) => (
                                          <tr key={index}>
                                            <td>{index + 1}</td>
                                            <td>{item.From}</td>
                                            <td>{item["From State"]}</td>
                                            <td>{item.To}</td>
                                            <td>{item["To State"]}</td>
                                            <td>{item.Commodity}</td>
                                            {/* <td>{item.Values}</td> */}
                                          </tr>
                                        ))}
                                      </tbody>
                                    </table>
                                  </div>
                                ) : (
                                  <p style={{ marginTop: 10 }}>
                                    No Rice routes
                                  </p>
                                )}
                              </div>
                            )}
                          </div>
                          <div></div>
                          {showMessage && (
                            <div>
                              {wheatData !== null && wheatData.length > 0 ? (
                                <div>
                                  <table>
                                    <thead>
                                      <tr style={{ margin: "auto" }}>
                                        <th
                                          style={{
                                            padding: "10px",
                                            width: "350px",
                                          }}
                                        >
                                          Sr. No
                                        </th>
                                        <th
                                          style={{
                                            padding: "10px",
                                            width: "350px",
                                          }}
                                        >
                                          Src RH
                                        </th>
                                        <th
                                          style={{
                                            padding: "10px",
                                            width: "350px",
                                          }}
                                        >
                                          Src state
                                        </th>
                                        <th
                                          style={{
                                            padding: "10px",
                                            width: "350px",
                                          }}
                                        >
                                          Dest RH
                                        </th>
                                        <th
                                          style={{
                                            padding: "10px",
                                            width: "350px",
                                          }}
                                        >
                                          Dest State
                                        </th>
                                        <th
                                          style={{
                                            padding: "10px",
                                            width: "350px",
                                          }}
                                        >
                                          commodity
                                        </th>
                                        {/* <th style={{ padding: "10px", width: "350px" }}>
                                        values
                                      </th> */}
                                      </tr>
                                    </thead>
                                    <tbody>
                                      {wheatData.map((item) => (
                                        <tr key={item["Unnamed: 0"]}>
                                          <td>{item["Unnamed: 0"] + 1}</td>
                                          <td>{item.From}</td>
                                          <td>{item["From State"]}</td>
                                          <td>{item.To}</td>
                                          <td>{item["To State"]}</td>
                                          <td>{item.Commodity}</td>
                                          {/* <td>{item.Values}</td> */}
                                        </tr>
                                      ))}
                                    </tbody>
                                  </table>
                                </div>
                              ) : (
                                <p>No Wheat routes</p>
                              )}
                            </div>
                          )}
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

                          <button
                            style={{ color: "white", marginLeft: "15px" }}
                            className="btn btn-danger dropdown-toggle"
                            onClick={viewGrid}
                          >
                            View Railhead Detailed Plan
                          </button>
                          {showMessage && (
                            <div style={{ marginTop: 15, marginLeft: 20 }}>
                              {riceData !== null && riceData.length > 0 ? (
                                <div>
                                  <table>
                                    <thead>
                                      <tr style={{ margin: "auto" }}>
                                        <th
                                          style={{
                                            padding: "10px",
                                            width: "200px",
                                          }}
                                        >
                                          Sr. No
                                        </th>
                                        <th
                                          style={{
                                            padding: "10px",
                                            width: "200px",
                                          }}
                                        >
                                          Src RH
                                        </th>
                                        <th
                                          style={{
                                            padding: "10px",
                                            width: "200px",
                                          }}
                                        >
                                          Src state
                                        </th>
                                        <th
                                          style={{
                                            padding: "10px",
                                            width: "200px",
                                          }}
                                        >
                                          Dest RH
                                        </th>
                                        <th
                                          style={{
                                            padding: "10px",
                                            width: "200px",
                                          }}
                                        >
                                          Dest state
                                        </th>
                                        <th
                                          style={{
                                            padding: "10px",
                                            width: "200px",
                                          }}
                                        >
                                          commodity
                                        </th>
                                        {/* <th style={{ padding: "10px", width: "350px" }}>
                                          values
                                        </th> */}
                                      </tr>
                                    </thead>
                                    <tbody>
                                      {riceData.map((item, index) => (
                                        <tr key={item["Unnamed: 0"]}>
                                          <td>{index + 1}</td>
                                          <td>{item.From}</td>
                                          <td>{item["From State"]}</td>
                                          <td>{item.To}</td>
                                          <td>{item["To State"]}</td>
                                          <td>{item.Commodity}</td>
                                          {/* <td>{item.Values}</td> */}
                                        </tr>
                                      ))}
                                    </tbody>
                                  </table>
                                </div>
                              ) : (
                                <p style={{ marginTop: 10 }}>No Rice routes</p>
                              )}
                            </div>
                          )}
                          {showMessage && (
                            <div style={{ marginTop: 15, marginLeft: 20 }}>
                              {wheatData !== null && wheatData.length > 0 ? (
                                <div>
                                  <table>
                                    <thead>
                                      <tr style={{ margin: "auto" }}>
                                        <th
                                          style={{
                                            padding: "10px",
                                            width: "200px",
                                          }}
                                        >
                                          Sr. No
                                        </th>
                                        <th
                                          style={{
                                            padding: "10px",
                                            width: "200px",
                                          }}
                                        >
                                          Src RH
                                        </th>
                                        <th
                                          style={{
                                            padding: "10px",
                                            width: "200px",
                                          }}
                                        >
                                          Src state
                                        </th>
                                        <th
                                          style={{
                                            padding: "10px",
                                            width: "200px",
                                          }}
                                        >
                                          Dest RH
                                        </th>
                                        <th
                                          style={{
                                            padding: "10px",
                                            width: "200px",
                                          }}
                                        >
                                          Dest State
                                        </th>
                                        <th
                                          style={{
                                            padding: "10px",
                                            width: "200px",
                                          }}
                                        >
                                          commodity
                                        </th>
                                        {/* <th style={{ padding: "10px", width: "350px" }}>
                                        values
                                      </th> */}
                                      </tr>
                                    </thead>
                                    <tbody>
                                      {wheatData.map((item, index) => (
                                        <tr key={index}>
                                          <td>{index + 1}</td>
                                          <td>{item.From}</td>
                                          <td>{item["From State"]}</td>
                                          <td>{item.To}</td>
                                          <td>{item["To State"]}</td>
                                          <td>{item.Commodity}</td>
                                          {/* <td>{item.Values}</td> */}
                                        </tr>
                                      ))}
                                    </tbody>
                                  </table>
                                </div>
                              ) : (
                                <p>No Wheat routes</p>
                              )}
                            </div>
                          )}
                        </div>
                      )}
                    </div>
                  )}
                  <br />
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
          </div>
        </div>
        <div
          style={{
            backgroundColor: "#ebab44b0",
            width: "20%",
            display: "flex",
            flexDirection: "column",
            alignContent: "center",
            alignItems: "center",
          }}
        >
          {/* <br /> */}
          {/* <div>
            <div class="progress yellow">
              <span class="progress-left">
                <span class="progress-bar"></span>
              </span>
              <span class="progress-right">
                <span class="progress-bar"></span>
              </span>
              <div class="progress-value">Steps</div>
            </div>
          </div> */}
          <span style={{ color: "black", fontSize: "32px" }}>Progress Bar</span>
          {number_check1 > 0 ||
          number_check2 > 0 ||
          supplyWeatCount > 0 ||
          destinationWheatCount > 0 ||
          progress.length > 0 ? (
            <div
              style={{
                padding: "8px 0",
                width: "90%",
                display: "flex",
                flexDirection: "column",
                border: "2px dashed black",
                marginTop: 15,
              }}
              id="console_"
            >
              <div
                style={{
                  margin: "0px 8px",
                  display: "flex",
                  flexDirection: "column",
                  gap: 8,
                }}
              >
                {number_check1 > 0 ? (
                  <div>{`Supply Value of Rice is ${number_check1}`}</div>
                ) : null}
                {number_check2 > 0 ? (
                  <div
                    style={{
                      color: number_check1 >= number_check2 ? "" : "red",
                    }}
                  >
                    {`Destination Value of Rice is ${number_check2}`}
                  </div>
                ) : null}
                {supplyWeatCount > 0 ? (
                  <div>{`Supply Value of Wheat is ${supplyWeatCount}`}</div>
                ) : null}
                {destinationWheatCount > 0 ? (
                  <div
                    style={{
                      color:
                        supplyWeatCount >= destinationWheatCount ? "" : "red",
                    }}
                  >
                    {`Destination Value of Wheat is ${destinationWheatCount}`}
                  </div>
                ) : null}
                {progress.length > 0 && progress.map((ele) => <div>{ele}</div>)}
                {isLoading ? (
                  <div
                    style={{
                      width: 200,
                      display: "grid",
                      gridTemplateColumns: "1fr auto",
                      alignItems: "end",
                    }}
                  >
                    Processing
                    <span
                      class="container"
                      style={{
                        display: "flex",
                        flexDirection: "row",
                        marginLeft: -10,
                        marginBottom: 4,
                        gap: 1,
                      }}
                    >
                      <div class="dot"></div>
                      <div class="dot"></div>
                      <div class="dot"></div>
                    </span>
                  </div>
                ) : null}
                {downloadMessage ? (
                  <div>
                    Solution has been done. Click on download RH to RH Detailed
                    plan
                  </div>
                ) : null}
              </div>
            </div>
          ) : null}
        </div>
      </div>
    </div>
  );
}

export default Daily_Planner;
