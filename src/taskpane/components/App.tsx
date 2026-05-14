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

  const [status, setStatus] = useState("ยังไม่ดำเนินการ");
  
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
        text: `${item.Job_x0020_details}`,
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

  // --- ปรับแต่ง ComboBox Styles เพื่อให้ Wrap Text หลังจากเลือกแล้ว ---
  const comboBoxStyles: Partial<IComboBoxStyles> = {
    root: { 
      width: '100%', 
      height: 'auto', // ให้ความสูงยืดหยุ่น
      minHeight: '32px' 
    },
    container: { 
      height: 'auto',
      overflow: 'visible' // สำคัญ: เพื่อให้ข้อความที่ยาวไม่ถูกตัดหายไป
    },
    input: {
      whiteSpace: 'normal', // อนุญาตให้ข้อความขึ้นบรรทัดใหม่
      wordBreak: 'break-word',
      height: 'auto',
      minHeight: '32px',
      lineHeight: '1.5',
      padding: '5px 0',
      overflow: 'visible'
    },
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

  const handleSubmit = async () => {
    if (!subject.trim()) {
      setSubjectError("กรุณากรอก Subject ก่อนบันทึก");
      return;
    }
    setIsSubmitting(true);
    const item = Office.context.mailbox.item;
    const submitUrl = "https://defaultb8d867c0b949455c95ddcee5324ed8.15.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/e902184684064f9f991e7ceb74a18807/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=o5Px4EMmRxBaqYQOhuicE40tsDBanclTxuR3M6uu9bg";
    const emailIdForFlow = Office.context.mailbox.convertToRestId(item.itemId, Office.MailboxEnums.RestVersion.v2_0);
  
    const payload = {
      Subject: subject,
      JobDetails: selectedJob.data.Job_x0020_details,
      JobType: selectedJob.data.Job_x0020_Type,
      AssignedTo: selectedOfficer.text,
      TrackingSLA: selectedJob.data.Tracking_x0020_SLA,
      Status: status,
      SendMail: selectedJob.data.Send_x0020_Email,
      ToEmail: selectedOfficer.data.Officer_Name?.Email || "",
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
        const result = await response.json();
        const finalCaseNo = result.caseNo || "COU_ERROR";
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
        <h2 style={{ color: '#0078d4', margin: '0 0 5px 0' }}>Case Assignment V 3.6</h2>
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

        <PrimaryButton 
          text={isSubmitting ? "กำลังบันทึก..." : "Submit Case"} 
          onClick={handleSubmit} 
          disabled={!selectedJob || !selectedOfficer || !subject.trim() || isSubmitting} 
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