import * as React from "react";
import * as ReactDOM from "react-dom";
import App from "./components/App";
import { initializeIcons } from "@fluentui/react";

/* global document, Office */

initializeIcons();

// ฟังก์ชันสำหรับ Render
const renderApp = () => {
  const container = document.getElementById("container");
  if (container) {
    ReactDOM.render(<App />, container);
  }
};

// ตรวจสอบว่า Office พร้อมไหม ถ้าเปิดใน Browser ปกติให้รันเลย
if (typeof Office !== 'undefined') {
  Office.onReady(() => {
    console.log("Office is ready!");
    renderApp();
  });
} else {
  // กรณีรันบน Browser ตรงๆ
  renderApp();
}