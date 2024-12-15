import React, { useState, useEffect } from "react";
import {
  FluentProvider,
  webLightTheme,
  Button,
  Input,
  Text,
  Switch,
} from "@fluentui/react-components";
import MarkdownCard from "./MarkdownCard";
import { Configuration, OpenAIApi } from "openai";
import OPENAI_API_KEY from "../../config/openaiKey";
import axios from "axios";
import ContentEditable from "react-contenteditable";

interface Frame2Props {
  switchToFrame3: (requestInput: string) => void;
  accessToken: string;
  requestInput: string;
}

interface ProfileData {
  customerProfile: string;
  objectName: string;
  folderName: string;
}

interface Email {
  outlookEmailId: string;
  customerProfile: string;
  rating: number;
}

const Frame2: React.FC<Frame2Props> = ({ switchToFrame3, accessToken, requestInput }) => {
  // State management
  const [propertyName, setPropertyName] = useState("Immobilie XXX");
  const [requestsInfo, setRequestsInfo] = useState("XXX der XXX Anfragen treffen auf die Profilbeschreibung zu");
  const [confirmationTemplate, setConfirmationTemplate] = useState("");
  const [rejectionTemplate, setRejectionTemplate] = useState("");
  const [customerProfile, setCustomerProfile] = useState("");
  const [emails, setEmails] = useState<Email[]>([]);
  const [deleteEmailsToggle, setDeleteEmailsToggle] = useState(
    localStorage.getItem("deleteEmailsToggle") === "true"
  );
  
  // Form validation states
  const [isFormValid, setIsFormValid] = useState(false);
  const [showConfirmationTemplateError, setShowConfirmationTemplateError] = useState(false);
  const [showRejectionTemplateError, setShowRejectionTemplateError] = useState(false);

  // Loading states
  const [isLoading, setIsLoading] = useState(true);
  const [profileData, setProfileData] = useState<ProfileData | null>(null);
  const [retryCount, setRetryCount] = useState(0);
  const MAX_RETRIES = 5;
  const RETRY_DELAY = 2000;

  const restId = Office.context.mailbox.item
    ? Office.context.mailbox.convertToRestId(
        Office.context.mailbox.item.itemId,
        Office.MailboxEnums.RestVersion.v2_0
      )
    : null;
  const emailId = restId;

  // API Functions
  const fetchProfileData = async (outlookEmailId: string): Promise<ProfileData | null> => {
    try {
      const encodedEmailId = encodeURIComponent(outlookEmailId);
      const baseUrl = 'https://cosmosdbbackendplugin.azurewebsites.net';
      
      const [profileResponse, nameResponse, folderResponse] = await Promise.all([
        fetch(`${baseUrl}/fetchCustomerProfile?outlookEmailId=${encodedEmailId}`),
        fetch(`${baseUrl}/fetchName?outlookEmailId=${encodedEmailId}`),
        fetch(`${baseUrl}/fetchFolderName?outlookEmailId=${encodedEmailId}`)
      ]);

      if (!profileResponse.ok || !nameResponse.ok || !folderResponse.ok) {
        throw new Error('One or more requests failed');
      }

      const [profileData, nameData, folderData] = await Promise.all([
        profileResponse.json(),
        nameResponse.json(),
        folderResponse.json()
      ]);

      return {
        customerProfile: profileData.customerProfile,
        objectName: nameData.objectname,
        folderName: folderData.folderName
      };
    } catch (error) {
      console.error('Error fetching profile data:', error);
      return null;
    }
  };

  const fetchEmailsByFolderName = async (folderName: string): Promise<Email[]> => {
    try {
      const encodedFolderName = encodeURIComponent(folderName);
      const response = await fetch(
        `https://cosmosdbbackendplugin.azurewebsites.net/fetchEmailsByFolderName?folderName=${encodedFolderName}`
      );
      const emails = await response.json();
      return emails;
    } catch (error) {
      console.error("Error fetching emails by folder name:", error);
      return [];
    }
  };

  // Folder Management Functions
  const ensureFolderExists = async (folderName: string): Promise<string | null> => {
    try {
      const response = await axios.get(
        "https://graph.microsoft.com/v1.0/me/mailFolders",
        {
          headers: {
            Authorization: `Bearer ${accessToken}`,
          },
        }
      );

      const folders = response.data.value;
      let folder = folders.find((f: any) => f.displayName === folderName);

      if (folder) {
        return folder.id;
      } else {
        const createFolderResponse = await axios.post(
          "https://graph.microsoft.com/v1.0/me/mailFolders",
          {
            displayName: folderName,
          },
          {
            headers: {
              Authorization: `Bearer ${accessToken}`,
              "Content-Type": "application/json",
            },
          }
        );
        return createFolderResponse.data.id;
      }
    } catch (error) {
      console.error(`Error ensuring folder '${folderName}' exists:`, error);
      return null;
    }
  };

  // Email Management Functions
// Replace the existing createDraftReplyAndMove function with this one:
const createDraftReplyAndMove = async () => {
  try {
    // Fetch latest data before processing
    const latestProfileData = await fetchProfileData(emailId);
    if (!latestProfileData?.folderName) {
      console.error("No folder name available");
      return;
    }

    // Fetch latest emails
    const latestEmails = await fetchEmailsByFolderName(latestProfileData.folderName);
    
    const numberOfAcceptedEmails = parseInt(requestInput, 10);
    if (isNaN(numberOfAcceptedEmails) || numberOfAcceptedEmails < 0) {
      console.error("Invalid number in requestInput");
      return;
    }

    // Sort emails by rating and split into accepted/rejected
    const sortedEmails = latestEmails.sort((a, b) => b.rating - a.rating);
    const acceptedEmails = sortedEmails
      .filter(email => email.rating > 7)
      .slice(0, numberOfAcceptedEmails);
    const rejectedEmails = sortedEmails
      .filter(email => email.rating <= 7)
      .concat(sortedEmails.slice(numberOfAcceptedEmails));

    // Create folders
    const acceptedFolderId = await ensureFolderExists("akzeptiert" + latestProfileData.folderName);
    const rejectedFolderId = await ensureFolderExists("abgelehnt" + latestProfileData.folderName);

    // Process emails
    const processEmails = async (emailList: Email[], folderId: string, template: string) => {
      for (const email of emailList) {
        await createDraftReplyForEmail(email.outlookEmailId, folderId, template);
        if (deleteEmailsToggle) {
          await deleteEmail(email.outlookEmailId);
        }
      }
    };

    await Promise.all([
      processEmails(acceptedEmails, acceptedFolderId!, confirmationTemplate),
      processEmails(rejectedEmails, rejectedFolderId!, rejectionTemplate)
    ]);

    switchToFrame3(requestInput);
  } catch (error) {
    console.error("Error in createDraftReplyAndMove:", error);
  }
};

  const createDraftReplyForEmail = async (emailId: string, folderId: string, template: string) => {
    try {
      const createDraftResponse = await axios.post(
        `https://graph.microsoft.com/v1.0/me/messages/${emailId}/createReply`,
        {
          comment: template,
        },
        {
          headers: {
            Authorization: `Bearer ${accessToken}`,
            "Content-Type": "application/json",
          },
        }
      );

      const draftMessageId = createDraftResponse.data.id;

      await axios.post(
        `https://graph.microsoft.com/v1.0/me/messages/${draftMessageId}/move`,
        {
          destinationId: folderId,
        },
        {
          headers: {
            Authorization: `Bearer ${accessToken}`,
            "Content-Type": "application/json",
          },
        }
      );

      console.log(`Draft reply for email ${emailId} created and moved.`);
    } catch (error) {
      console.error(`Error creating draft reply for email ${emailId}:`, error);
    }
  };

  const deleteEmail = async (emailId: string) => {
    try {
      await axios.delete(
        `https://graph.microsoft.com/v1.0/me/messages/${emailId}`,
        {
          headers: {
            Authorization: `Bearer ${accessToken}`,
          },
        }
      );
      console.log(`Email ${emailId} deleted from inbox.`);
    } catch (error) {
      console.error(`Error deleting email ${emailId}:`, error);
    }
  };

  // UI Event Handlers
  const toggleDeleteEmails = () => {
    const newToggle = !deleteEmailsToggle;
    setDeleteEmailsToggle(newToggle);
    localStorage.setItem("deleteEmailsToggle", newToggle.toString());
    console.log(`Delete emails toggle set to: ${newToggle}`);

  };

  const handleCardClick = (email: Email) => {
    openEmailItem(email.outlookEmailId);
  };

  const openEmailItem = (itemId: string) => {
    try {
      const restId = Office.context.mailbox.convertToRestId(
        itemId,
        Office.MailboxEnums.RestVersion.v1_0
      );
      Office.context.mailbox.displayMessageForm(restId);
    } catch (error) {
      console.error("Error opening email item:", error);
    }
  };

  // Effects
  useEffect(() => {
    let timeoutId: NodeJS.Timeout;
    let isMounted = true;

    const attemptFetch = async () => {
      if (!emailId || !isMounted) return;

      const data = await fetchProfileData(emailId);
      
      if (!isMounted) return;

      if (data) {
        setProfileData(data);
        setCustomerProfile(data.customerProfile);
        setPropertyName(data.objectName);
        
        // Fetch and process emails
        const fetchedEmails = await fetchEmailsByFolderName(data.folderName);
        const sortedEmails = fetchedEmails.sort((a, b) => b.rating - a.rating);
        setEmails(sortedEmails);
        
        // Update requests info
        const numberOfAcceptedEmails = parseInt(requestInput, 10) || 0;
        setRequestsInfo(`${numberOfAcceptedEmails} der ${sortedEmails.length} Anfragen treffen auf die Profilbeschreibung zu`);
        
        setIsLoading(false);
        setRetryCount(0);
      } else if (retryCount < MAX_RETRIES) {
        timeoutId = setTimeout(() => {
          if (isMounted) {
            setRetryCount(prev => prev + 1);
          }
        }, RETRY_DELAY);
      } else {
        setIsLoading(false);
        console.error('Failed to fetch profile data after maximum retries');
      }
    };

    if (isLoading && retryCount < MAX_RETRIES) {
      attemptFetch();
    }

    return () => {
      isMounted = false;
      clearTimeout(timeoutId);
    };
  }, [emailId, retryCount, isLoading, requestInput]);

  useEffect(() => {
    const validateForm = () => {
      const isConfirmationTemplateValid = confirmationTemplate.trim() !== "";
      const isRejectionTemplateValid = rejectionTemplate.trim() !== "";

      setShowConfirmationTemplateError(!isConfirmationTemplateValid);
      setShowRejectionTemplateError(!isRejectionTemplateValid);

      setIsFormValid(isConfirmationTemplateValid && isRejectionTemplateValid);
    };

    validateForm();
  }, [confirmationTemplate, rejectionTemplate]);

  // Loading UI effect
  useEffect(() => {
    if (isLoading) {
      setPropertyName("Lädt...");
      setCustomerProfile("Lädt...");
    }
  }, [isLoading]);

  return (
    <FluentProvider theme={webLightTheme}>
      <div style={{ padding: "20px", margin: "0 auto" }}>
        {/* Property Information */}
        <MarkdownCard markdown={`**${propertyName}**`} />
        <MarkdownCard markdown={requestsInfo} />
        <MarkdownCard markdown={customerProfile} />

        

        {/* Templates */}
        <ContentEditable
          html={confirmationTemplate}
          onChange={(e) => setConfirmationTemplate(e.target.value)}
          tagName="div"
          style={{
            marginBottom: "10px",
            height: '100px',
            border: '1px solid #ccc',
            padding: '10px',
            overflow: 'auto',
            whiteSpace: 'pre-wrap',
            wordWrap: 'break-word',
            resize: 'both',
            boxSizing: 'border-box',
          }}
          
        />
        {showConfirmationTemplateError && <Text style={{ color: "red" }}>Bestätigungsemail-Template ist erforderlich</Text>}
        <ContentEditable
          html={rejectionTemplate}
          onChange={(e) => setRejectionTemplate(e.target.value)}
          tagName="div"
          style={{
            marginBottom: "10px",
            height: '100px',
            border: '1px solid #ccc',
            padding: '10px',
            overflow: 'auto',
            whiteSpace: 'pre-wrap',
            wordWrap: 'break-word',
            resize: 'both',
            boxSizing: 'border-box',
          }}
          
        />
        {showRejectionTemplateError && <Text style={{ color: "red" }}>Absageemail-Template ist erforderlich</Text>}
        {/* Toggle Switch */}
        <Switch
          checked={deleteEmailsToggle}
          onChange={toggleDeleteEmails}
          label="Delete Emails from Inbox"
          style={{ width: "100%", marginTop: "10px" }}
        />
        {/* Drafts Button */}
        <Button
          appearance="primary"
          style={{ width: "100%" }}
          onClick={createDraftReplyAndMove}
          disabled={!isFormValid}
        >
          Drafts erstellen
        </Button>
      </div>
      {/* Display sorted customer profiles */}
      {emails.map((email, index) => (
          <div key={index} onClick={() => handleCardClick(email)}>
            <MarkdownCard markdown={email.customerProfile} />
          </div>
        ))}

    </FluentProvider>
  );
};

export default Frame2;
