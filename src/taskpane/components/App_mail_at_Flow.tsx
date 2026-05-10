import * as React from "react";
import { useState, useEffect } from "react";
import { 
  PrimaryButton, 
  TextField, 
  Dropdown, 
  IDropdownOption, 
  Stack, 
  Dialog, 
  DialogType, 
  DialogFooter 
} from "@fluentui/react";

/* global Office */

const App: React.FC = () => {
  const [caseNo, setCaseNo] = useState("");
  const [subject, setSubject] = useState("");
  const [subjectError, setSubjectError] = useState<string | undefined>(undefined); // สำหรับ Validate
  const [jobOptions, setJobOptions] = useState<IDropdownOption[]>([]);
  const [selectedJob, setSelectedJob] = useState<any>(null);
  const [status, setStatus] = useState("ยังไม่ดำเนินการ");
  const [isDialogOpen, setIsDialogOpen] = useState(false);
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [dialogData, setDialogData] = useState({ title: "", message: "" });

  useEffect(() => {
    Office.onReady(async () => {
      if (Office.context.mailbox.item) {
        const initialSubject = Office.context.mailbox.item.subject || "";
        setSubject(initialSubject);
        // เช็คตั้งแต่วินาทีแรกที่เปิด ถ้าไม่มี Subject ให้ขึ้น Error เลย
        if (!initialSubject.trim()) {
          setSubjectError("กรุณากรอก Subject ก่อนบันทึก");
        }
      }
      const date = new Date().toISOString().slice(0, 10).replace(/-/g, "");
      const randomNum = Math.floor(1000 + Math.random() * 9000);
      setCaseNo(`COU_${date}_${randomNum}`);
      fetchJobDetails();
    });
  }, []);

  const openMsg = (title: string, msg: string) => {
    setDialogData({ title: title, message: msg });
    setIsDialogOpen(true);
  };

  const fetchJobDetails = async () => {
    const getMasterUrl = "https://defaultb8d867c0b949455c95ddcee5324ed8.15.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/1ba262af176243dc8e39d82233fc6bd7/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=zldU17I_mV3UN-FK3ASnPIOKcuMT8Zvpn9KjFtEm114"; 
    try {
      const response = await fetch(getMasterUrl); 
      if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
      const data = await response.json();
      const items = Array.isArray(data) ? data : (data.value || []);
      const options = items.map((item: any) => ({
        key: item.ID,
        text: `${item.ID}. ${item.Job_x0020_details} (${item.AssignedTo || 'ไม่ระบุคนรับ'})`,
        data: item 
      }));
      setJobOptions(options);
    } catch (e) {
      console.error("Fetch Master Error:", e);
    }
  };

  const handleSubmit = async () => {
    // Double Check Validation
    if (!subject.trim()) {
      setSubjectError("กรุณากรอก Subject ก่อนบันทึก");
      return;
    }

    setIsSubmitting(true);
    const submitUrl = "https://defaultb8d867c0b949455c95ddcee5324ed8.15.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/e902184684064f9f991e7ceb74a18807/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=o5Px4EMmRxBaqYQOhuicE40tsDBanclTxuR3M6uu9bg";
    const item = Office.context.mailbox.item;
    const emailIdForFlow = Office.context.mailbox.convertToRestId(item.itemId, Office.MailboxEnums.RestVersion.v2_0);
    
    const payload = {
      CaseNo: caseNo,
      Subject: subject,
      JobDetails: selectedJob.data.Job_x0020_details,
      JobType: selectedJob.data.Job_x0020_Type,
      AssignedTo: selectedJob.data.AssignedTo,
      TrackingSLA: selectedJob.data.Tracking_x0020_SLA,
      Status: status,
      SendMail: selectedJob.data.Send_x0020_Email,
      ToEmail: selectedJob.data.AssignedToEmail,
      ReceiveDatetime: item.dateTimeCreated.toISOString(),
      EmailUrl: `https://outlook.office.com/mail/deeplink/read/${encodeURIComponent(item.itemId)}`,
      EmailID: emailIdForFlow
    };

    try {
      const response = await fetch(submitUrl, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
      });

      if (response.status === 200) {
        // กรณีสำเร็จ 100%
        openMsg("บันทึกสำเร็จ", `บันทึกเคส ${caseNo} เรียบร้อยแล้ว`);
      } 
      else if (response.status === 201) {
        // กรณีบันทึกสำเร็จ แต่ Flow มีบาง Step พัง (เช่น Reply หรือ Category)
        const errorData = await response.json();
        openMsg("บันทึกสำเร็จ (พบปัญหาบางขั้นตอน)", errorData.message || "มีบางขั้นตอนในระบบทำงานไม่สมบูรณ์");
        setIsSubmitting(false); // ปลดล็อกปุ่มเพื่อให้ตรวจสอบหรือกดซ้ำได้ถ้าจำเป็น
      } 
      else {
        openMsg("เกิดข้อผิดพลาด", `ระบบตอบกลับด้วยรหัส: ${response.status}`);
        setIsSubmitting(false);
      }
    } catch (error) {
      console.error("Network Error:", error);
      openMsg("ข้อผิดพลาดเครือข่าย", "ไม่สามารถติดต่อ Server ได้ (อาจเกิดจากปัญหา CORS หรือ Internet)");
      setIsSubmitting(false);
    }
  };

  const closeAddin = () => {
    Office.context.ui.closeContainer();
  };

  return (
    <div style={{ padding: 20 }}>
      <Stack tokens={{ childrenGap: 15 }}>
        <h2 style={{ color: '#0078d4', margin: '0 0 10px 0' }}>Case Assignment V 2.3</h2>
        
        <TextField label="Case No:" value={caseNo} readOnly />
        
        <TextField 
          label="Subject:" 
          value={subject} 
          required
          errorMessage={subjectError}
          onChange={(_, v) => {
            setSubject(v || "");
            if (v && v.trim() !== "") {
              setSubjectError(undefined);
            } else {
              setSubjectError("กรุณากรอก Subject ก่อนบันทึก");
            }
          }} 
        />

        <Dropdown
          label="Assign To (Job Details):"
          placeholder="เลือกรายการจาก Job Details"
          options={jobOptions}
          onChange={(_, opt) => setSelectedJob(opt)}
          styles={{ root: { width: '100%' } }}
        />

        <Dropdown
          label="Status:"
          defaultSelectedKey="ยังไม่ดำเนินการ"
          options={[{key:'ยังไม่ดำเนินการ', text:'ยังไม่ดำเนินการ'}, {key:'ปิดเคส', text:'ปิดเคส'}]}
          onChange={(_, opt) => setStatus(opt?.key as string)}
        />

        <PrimaryButton 
          text={isSubmitting ? "กำลังบันทึก..." : "Submit Case"} 
          onClick={handleSubmit} 
          // ปุ่มจะกดยังไม่ได้ถ้าเลือกงานไม่ครบ หรือ Subject ว่าง
          disabled={!selectedJob || !subject.trim() || isSubmitting} 
        />
      </Stack>

      <Dialog
        hidden={!isDialogOpen}
        onDismiss={closeAddin}
        dialogContentProps={{
          type: DialogType.normal,
          title: dialogData.title,
          subText: dialogData.message
        }}
        modalProps={{ isBlocking: true }}
      >
        <DialogFooter>
          <PrimaryButton onClick={closeAddin} text="ตกลง" />
        </DialogFooter>
      </Dialog>
    </div>
  );
};

export default App;