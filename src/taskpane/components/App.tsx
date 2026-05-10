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
  IDropdownStyles
} from "@fluentui/react";

/* global Office */

const App: React.FC = () => {
  const [subject, setSubject] = useState("");
  const [subjectError, setSubjectError] = useState<string | undefined>(undefined);
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
        if (!initialSubject.trim()) setSubjectError("กรุณากรอก Subject ก่อนบันทึก");
      }
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
      const data = await response.json();
      const items = Array.isArray(data) ? data : (data.value || []);
      const options = items.map((item: any) => ({
        key: item.ID,
        text: `${item.ID}. ${item.Job_x0020_details} (${item.AssignedTo || 'ไม่ระบุคนรับ'})`,
        data: item 
      }));
      setJobOptions(options);
    } catch (e) {
      console.error("Fetch Error:", e);
    }
  };

  const dropdownStyles: Partial<IDropdownStyles> = {
    dropdownItem: {
      whiteSpace: 'normal',
      height: 'auto',
      lineHeight: '1.4',
      padding: '8px 12px',
      borderBottom: '1px solid #eee'
    },
    title: {
      height: 'auto',
      minHeight: '32px',
      lineHeight: '1.4',
      padding: '5px 12px'
    }
  };

  const onRenderOption = (option?: IDropdownOption): JSX.Element => {
    return (
      <div title={option?.text} style={{ wordWrap: 'break-word', width: '100%' }}>
        {option?.text}
      </div>
    );
  };

  const assignCategory = (categoryName: string) => {
    if (Office.context.mailbox.item.categories) {
      Office.context.mailbox.item.categories.addAsync([categoryName]);
    }
  };

  const assignCategories = (categoryNames: string[]) => {
    const item = Office.context.mailbox.item;
    if (item && item.categories) {
      const cleanCategories = categoryNames
        .filter(name => name && name.trim() !== "")
        .map(name => name.trim());

      cleanCategories.forEach(cat => {
        item.categories.addAsync([cat], (result) => {
          if (result.status === Office.AsyncResultStatus.Failed) {
            console.warn(`[Skip] Category "${cat}" not found in Master List. Error: ${result.error.code}`);
          } else {
            console.log(`[Success] Assigned: ${cat}`);
          }
        });
      });
    }
  };
  const forwardViaFlow = async (emailId: string, bodyHtml: string, toEmail: string) => {
    const forwardFlowUrl = "https://defaultb8d867c0b949455c95ddcee5324ed8.15.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/38527c197e144e6e8fc25075c7005f69/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=Ux3PJS3kd_qcZer8pV4ys9g3-YFtPUGNK5qPRK4EMQs";
    
    try {
      await fetch(forwardFlowUrl, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          emailId: emailId,
          body: bodyHtml,
          to: toEmail
        })
      });
      console.log("Forward via Flow success");
    } catch (error) {
      console.error("Forward Flow Error:", error);
    }
  };
	const openInternalReply = (finalCaseNo: string) => {
		const item = Office.context.mailbox.item;
		const bodyHtml = `
		  <div style="font-family: Calibri, sans-serif; font-size: 11pt;">
			เรียน ทีมงานที่เกี่ยวข้อง,<br/><br/>
			บันทึกเคสเรียบร้อยแล้ว:<br/>
			<b>เลขที่เคส:</b> ${finalCaseNo}<br/>
			<b>เรื่อง:</b> ${subject}<br/>
			<b>รายละเอียดงาน:</b> ${selectedJob?.data?.Job_x0020_details || ""}<br/>
			<b>ผู้รับผิดชอบ:</b> ${selectedJob?.data?.AssignedTo || ""}<br/><br/>
			ขอบคุณครับ
		  </div>
		`;

		const shouldForward = selectedJob?.data?.Send_x0020_Email === true || selectedJob?.data?.Send_x0020_Email === "true";
		const toEmail = selectedJob?.data?.AssignedToEmail || "";
		const emailIdForFlow = Office.context.mailbox.convertToRestId(item.itemId, Office.MailboxEnums.RestVersion.v2_0);
		if (shouldForward) {
			// เรียกใช้ Flow แทนการเปิด Form
			forwardViaFlow(emailIdForFlow, bodyHtml, toEmail);
			// ไม่ต้องสั่งเปิด Form ใดๆ เพราะ Flow จะส่งเมลให้เองหลังบ้าน
				  
		} else {
		  // สำหรับ Reply ใช้แบบนี้จะปลอดภัยกว่าในเวอร์ชันเก่า
		  item.displayReplyForm({
			htmlBody: bodyHtml,
			attachments: [] 
		  });
		}
	  };

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
        const result = await response.json();
        const finalCaseNo = result.caseNo || "COU_ERROR";
        
        let categoriesToApply = ["บันทึกเคส-COU-เรียบร้อย"];
        const assignedValue = selectedJob?.data?.AssignedTo;
        
        if (assignedValue && assignedValue.toString().trim() !== "") {
          const splitNames = assignedValue.toString().split(",").map(name => name.trim());
          categoriesToApply = [...categoriesToApply, ...splitNames];
        }

        console.log("DEBUG [Split Result] Final array to assign:", categoriesToApply);

        assignCategories(categoriesToApply);
        openInternalReply(finalCaseNo);
        openMsg("สำเร็จ", `บันทึกเคสเลขที่ ${finalCaseNo} เรียบร้อยแล้ว`);
      } else {
        openMsg("เกิดข้อผิดพลาด", `Server ตอบกลับด้วย Status: ${response.status}`);
        setIsSubmitting(false);
      }
    } catch (error) {
      console.error("Submit Error:", error);
      openMsg("ข้อผิดพลาดเครือข่าย", "ไม่สามารถติดต่อ Server ได้ กรุณาตรวจสอบการตั้งค่า JSON ใน Flow");
      setIsSubmitting(false);
    }
  };

  const closeAddin = () => Office.context.ui.closeContainer();

  return (
    <div style={{ padding: '10px 20px' }}>
      <Stack tokens={{ childrenGap: 15 }}>
        <h2 style={{ color: '#0078d4', margin: '0 0 5px 0' }}>Case Assignment V 2.9</h2>
        
        <TextField 
          label="Subject:" 
          value={subject} 
          required
          errorMessage={subjectError}
          onChange={(_, v) => {
            setSubject(v || "");
            setSubjectError(v?.trim() ? undefined : "กรุณากรอก Subject ก่อนบันทึก");
          }} 
        />

        <Dropdown
          label="Assign To (Job Details):"
          placeholder="เลือกรายการงาน"
          options={jobOptions}
          styles={dropdownStyles}
          onRenderOption={onRenderOption}
          onChange={(_, opt) => setSelectedJob(opt)}
        />

        <Dropdown
          label="Status:"
          defaultSelectedKey="ยังไม่ดำเนินการ"
          options={[
            { key: 'ยังไม่ดำเนินการ', text: 'ยังไม่ดำเนินการ' },
            { key: 'ปิดเคส', text: 'ปิดเคส' }
          ]}
          onChange={(_, opt) => setStatus(opt?.key as string)}
        />

        <PrimaryButton 
          text={isSubmitting ? "กำลังบันทึก..." : "Submit Case"} 
          onClick={handleSubmit} 
          disabled={!selectedJob || !subject.trim() || isSubmitting} 
          styles={{ root: { marginTop: 10 } }}
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