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
  DialogFooter,
  IDropdownStyles,
  ComboBox,
  IComboBoxOption,
  IComboBoxStyles,
  IComboBox
} from "@fluentui/react";

/* global Office */

const App: React.FC = () => {
  const [subject, setSubject] = useState("");
  const [subjectError, setSubjectError] = useState<string | undefined>(undefined);
  const [jobOptions, setJobOptions] = useState<IComboBoxOption[]>([]);
  const [selectedJob, setSelectedJob] = useState<any>(null);
  
  const [allOfficers, setAllOfficers] = useState<any[]>([]); 
  const [officerOptions, setOfficerOptions] = useState<IDropdownOption[]>([]); 
  const [selectedOfficer, setSelectedOfficer] = useState<any>(null);

  // State สำหรับผู้ทำรายการ (Submitter)
  const [selectedSubmitter, setSelectedSubmitter] = useState<IDropdownOption | null>(null);

  const [status, setStatus] = useState("ยังไม่ดำเนินการ");
  const [remarks, setRemarks] = useState(""); // เพิ่ม State สำหรับ Remarks
  const [isDialogOpen, setIsDialogOpen] = useState(false);
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [dialogData, setDialogData] = useState({ title: "", message: "" });

  useEffect(() => {
    Office.onReady(async () => {
      if (Office.context.mailbox.item) {
        const initialSubject = Office.context.mailbox.item.subject || "";
        setSubject(initialSubject);
        if (!initialSubject.trim()) setSubjectError("กรุณากรอก Subject ก่อนบันทึก");
      }
      fetchJobDetails();
      fetchAllOfficers(); 
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
      const data = await response.json();
      const items = Array.isArray(data) ? data : (data.value || []);
      const options = items.map((item: any) => ({
        key: item.ID.toString(),
        text: `${item.Job_x0020_details} (${item.Primary_Assign})`,
        data: item 
      }));
      setJobOptions(options);
    } catch (e) {
      console.error("Fetch Job Error:", e);
    }
  };

  const fetchAllOfficers = async () => {
    const officerListUrl = "https://defaultb8d867c0b949455c95ddcee5324ed8.15.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/1888383f8d5a435e8e5ae76e27d7b501/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=e8WWiidIKUsj-ULKT3m4Qye0ZSFYybqtKVEYZCXDdxU"; 
    try {
      const response = await fetch(officerListUrl);
      const data = await response.json();
      const items = Array.isArray(data) ? data : (data.value || []);
      setAllOfficers(items);
      
      const options = items.map((item: any) => ({
        key: item.Title, 
        text: item.Title,
        data: item 
      }));

      // Logic: ดึงค่า Submitter ที่เคยเลือกไว้จาก localStorage
      const savedSubmitterKey = localStorage.getItem("savedSubmitterKey");
      if (savedSubmitterKey) {
        const found = options.find(opt => opt.key === savedSubmitterKey);
        if (found) setSelectedSubmitter(found);
      }
    } catch (e) {
      console.error("Fetch All Officers Error:", e);
    }
  };

  const filterOfficerList = (selectedJobId: string, isBackupAll: boolean) => {
    let filtered: any[] = [];
    if (isBackupAll) {
      filtered = allOfficers;
    } else {
      filtered = allOfficers.filter((officer: any) => {
        const jobIdsArray = officer["jobid#Id"] || [];
        if (jobIdsArray.length > 0) {
          return jobIdsArray.some((id: any) => id.toString() === selectedJobId);
        }
        const jobIdsObjects = officer.jobid || [];
        return jobIdsObjects.some((j: any) => j && j.Id && j.Id.toString() === selectedJobId);
      });
    }

    const options = filtered.map((item: any) => ({
      key: item.Title, 
      text: item.Title,
      data: item 
    }));

    setOfficerOptions(options);
    return options; 
  };

  const comboBoxStyles: Partial<IComboBoxStyles> = {
    root: { width: '100%', height: 'auto', minHeight: '32px' },
    container: { height: 'auto' },
    input: { whiteSpace: 'normal', height: 'auto', minHeight: '32px', lineHeight: '1.5', padding: '5px 0' },
    callout: { maxWidth: '300px' },
    optionsContainer: { maxHeight: 400 },
  };

  const dropdownStyles: Partial<IDropdownStyles> = {
    dropdownItem: { whiteSpace: 'normal', height: 'auto', lineHeight: '1.4', padding: '8px 12px', borderBottom: '1px solid #eee' },
    title: { height: 'auto', minHeight: '32px', lineHeight: '1.4', padding: '5px 12px', whiteSpace: 'normal' }
  };

  const onRenderComboBoxOption = (option?: IComboBoxOption): JSX.Element => (
    <div style={{ whiteSpace: 'normal', wordWrap: 'break-word', padding: '4px 0', lineHeight: '1.4' }}>
      {option?.text}
    </div>
  );

  const assignCategories = (categoryNames: string[]) => {
    return new Promise<void>((resolve) => {
      const item = Office.context.mailbox.item;
      if (item && item.categories) {
        const cleanCategories = categoryNames.filter(name => name && name.trim() !== "").map(name => name.trim());
        item.categories.addAsync(cleanCategories, (asyncResult) => {
          if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
            console.error("Failed to assign categories:", asyncResult.error.message);
          }
          resolve();
        });
      } else { resolve(); }
    });
  };

  const forwardViaFlow = async (emailId: string, bodyHtml: string, toEmail: string, shouldForward: boolean) => {
    const forwardFlowUrl = "https://defaultb8d867c0b949455c95ddcee5324ed8.15.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/38527c197e144e6e8fc25075c7005f69/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=Ux3PJS3kd_qcZer8pV4ys9g3-YFtPUGNK5qPRK4EMQs";
    try {
      await fetch(forwardFlowUrl, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ emailId, body: bodyHtml, to: toEmail, isForward: shouldForward })
      });
    } catch (error) { console.error("Forward Flow Error:", error); }
  };

  const openInternalReply = (finalCaseNo: string) => {
    const item = Office.context.mailbox.item;
    const remarksHtml = remarks && remarks.trim() !== "" ? `<b>หมายเหตุ:</b> ${remarks}<br/>` : "";
    const bodyHtml = `
      <div style="font-family: Calibri, sans-serif; font-size: 11pt;">
        เรียน ทีมงานที่เกี่ยวข้อง,<br/><br/>
        บันทึกเคสเรียบร้อยแล้ว:<br/>
        <b>เลขที่เคส:</b> ${finalCaseNo}<br/>
        <b>เรื่อง:</b> ${subject}<br/>
        <b>รายละเอียดงาน:</b> ${selectedJob?.data?.Job_x0020_details || ""}<br/>
        <b>ผู้รับผิดชอบ:</b> ${selectedOfficer?.text || ""}<br/>
        ${remarksHtml}<br/>
        ขอบคุณครับ
      </div>
    `;
    const shouldForward = selectedJob?.data?.Send_x0020_Email === true || selectedJob?.data?.Send_x0020_Email === "true";
    const toEmail = selectedOfficer?.data?.Officer_Name?.Email || ""; 
    const emailIdForFlow = Office.context.mailbox.convertToRestId(item.itemId, Office.MailboxEnums.RestVersion.v2_0);
    forwardViaFlow(emailIdForFlow, bodyHtml, toEmail, shouldForward);
  };

  const handleSubmit = async () => {
    if (!subject.trim() || !selectedSubmitter) {
      openMsg("คำเตือน", "กรุณากรอกข้อมูลและเลือกผู้บันทึกงานให้ครบถ้วน");
      return;
    }
    setIsSubmitting(true);
    const item = Office.context.mailbox.item;
    const submitUrl = "https://defaultb8d867c0b949455c95ddcee5324ed8.15.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/e902184684064f9f991e7ceb74a18807/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=o5Px4EMmRxBaqYQOhuicE40tsDBanclTxuR3M6uu9bg";
    const emailIdForFlow = Office.context.mailbox.convertToRestId(item.itemId, Office.MailboxEnums.RestVersion.v2_0);
  
    const payload = {
      Subject: subject,
      JobDetails: selectedJob?.data?.Job_x0020_details,
      JobType: selectedJob?.data?.Job_x0020_Type,
      AssignedTo: selectedOfficer?.text,
      TrackingSLA: selectedJob?.data?.Tracking_x0020_SLA,
      Status: status,
      Remarks: remarks, // ส่งค่า Remarks เพิ่มเข้าไปใน Payload
      SendMail: selectedJob?.data?.Send_x0020_Email,
      ToEmail: selectedOfficer?.data?.Officer_Name?.Email || "",
      ReceiveDatetime: item.dateTimeCreated.toISOString(),
      EmailUrl: `https://outlook.office.com/mail/deeplink/read/${encodeURIComponent(item.itemId)}`,
      EmailID: emailIdForFlow,
      Submit_user: selectedSubmitter.text // ส่งชื่อผู้ทำรายการไปเก็บที่ SharePoint
    };

    try {
      const response = await fetch(submitUrl, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
      });
      if (response.status === 200) {
        const result = await response.json();
        const finalCaseNo = result.caseNo || "COU_ERROR";
        await assignCategories(["บันทึกเคส-COU-เรียบร้อย", selectedOfficer?.text]);
        openInternalReply(finalCaseNo);
        openMsg("สำเร็จ", `บันทึกเคสเลขที่ ${finalCaseNo} เรียบร้อยแล้ว`);
      } else {
        openMsg("เกิดข้อผิดพลาด", `Status: ${response.status}`);
        setIsSubmitting(false);
      }
    } catch (error) {
      console.error("Submit Error:", error);
      setIsSubmitting(false);
    }
  };

  return (
    <div style={{ padding: '10px 20px' }}>
      <Stack tokens={{ childrenGap: 15 }}>
        <h2 style={{ color: '#0078d4', margin: '0 0 5px 0' }}>Case Assignment V 3.9</h2>
        
        <Dropdown
          label="ผู้บันทึกงาน (Submitter):"
          placeholder="เลือกชื่อของคุณ"
          options={allOfficers.map(item => ({ key: item.Title, text: item.Title }))}
          selectedKey={selectedSubmitter ? selectedSubmitter.key : undefined}
          onChange={(_, opt) => {
            if (opt) {
              setSelectedSubmitter(opt);
              localStorage.setItem("savedSubmitterKey", opt.key as string); // จำค่าไว้ในเครื่อง
            }
          }}
          styles={dropdownStyles}
        />

        <TextField label="Subject:" value={subject} onChange={(_, v) => setSubject(v || "")} />

        <ComboBox
          label="Assign To (Job Details):"
          placeholder="พิมพ์เพื่อค้นหา หรือเลือกงาน"
          options={jobOptions}
          allowFreeform={true}
          autoComplete="on"
          styles={comboBoxStyles}
          onRenderOption={onRenderComboBoxOption}
          selectedKey={selectedJob ? selectedJob.key : undefined}
          onChange={(_, opt) => {
            if (opt) {
              setSelectedJob(opt);
              const isBackupAll = opt.data?.Backup_all === true || opt.data?.Backup_all === "true";
              const filteredOptions = filterOfficerList(opt.key.toString(), isBackupAll);
              
              const primaryAssignName = opt.data?.Primary_Assign;
              if (primaryAssignName) {
                const defaultOfficer = filteredOptions.find(off => off.text === primaryAssignName);
                if (defaultOfficer) {
                  setSelectedOfficer(defaultOfficer);
                } else { setSelectedOfficer(null); }
              } else { setSelectedOfficer(null); }
            } else {
              setSelectedJob(null);
              setOfficerOptions([]);
              setSelectedOfficer(null);
            }
          }}
        />

        <Dropdown
          label="เลือก Assign To:"
          placeholder="เลือกเจ้าหน้าที่"
          options={officerOptions}
          selectedKey={selectedOfficer ? selectedOfficer.key : undefined}
          styles={dropdownStyles}
          onChange={(_, opt) => setSelectedOfficer(opt)}
          disabled={!selectedJob}
        />

        <Dropdown
          label="Status:"
          selectedKey={status} 
          options={[
            { key: 'ยังไม่ดำเนินการ', text: 'ยังไม่ดำเนินการ' }, 
            { key: 'ปิดเคส', text: 'ปิดเคส' }
          ]}
          onChange={(_, opt) => { if (opt) setStatus(opt.key as string); }}
        />

        <TextField 
          label="Remarks:" 
          multiline 
          rows={2} 
          value={remarks} 
          onChange={(_, v) => setRemarks(v || "")} 
        />

        <PrimaryButton 
          text={isSubmitting ? "กำลังบันทึก..." : "Submit Case"} 
          onClick={handleSubmit} 
          disabled={!selectedJob || !selectedOfficer || !subject.trim() || !selectedSubmitter || isSubmitting} 
          styles={{ root: { marginTop: 10 } }}
        />
      </Stack>

      <Dialog
        hidden={!isDialogOpen}
        onDismiss={() => Office.context.ui.closeContainer()}
        dialogContentProps={{ type: DialogType.normal, title: dialogData.title, subText: dialogData.message }}
        modalProps={{ isBlocking: true }}
      >
        <DialogFooter>
          <PrimaryButton onClick={() => Office.context.ui.closeContainer()} text="ตกลง" />
        </DialogFooter>
      </Dialog>
    </div>
  );
};

export default App;