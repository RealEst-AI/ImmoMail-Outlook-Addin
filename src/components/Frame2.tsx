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
  switchToFrame3: (requestInput) => void;
  accessToken: string;
  requestInput: string;
}

const Frame2: React.FC<Frame2Props> = ({ switchToFrame3, accessToken, requestInput }) => {
  // State for the dynamic values
  const [propertyName, setPropertyName] = useState("Immobilie XXX");
  const [requestsInfo, setRequestsInfo] = useState("XXX der XXX Anfragen treffen auf die Profilbeschreibung zu");
  const [confirmationTemplate, setConfirmationTemplate] = useState("");
  const [rejectionTemplate, setRejectionTemplate] = useState("");
  const [customerProfile, setCustomerProfile] = useState("");
  const [emails, setEmails] = useState([]); // State for emails
  const [deleteEmailsToggle, setDeleteEmailsToggle] = useState(localStorage.getItem("deleteEmailsToggle") === "true");

  const [isFormValid, setIsFormValid] = useState(false);
  const [showConfirmationTemplateError, setShowConfirmationTemplateError] = useState(false);
  const [showRejectionTemplateError, setShowRejectionTemplateError] = useState(false);

  const restId = Office.context.mailbox.item ? Office.context.mailbox.convertToRestId(
    Office.context.mailbox.item.itemId,
    Office.MailboxEnums.RestVersion.v2_0
  ) : null;
  console.log("REST-formatted Item ID:", restId);
  const emailId =restId;
  const fetchCustomerProfileFromBackend = async (outlookEmailId: string) => {
    try {
      const encodedEmailId = encodeURIComponent(outlookEmailId);
      const response = await fetch(
        `https://cosmosdbbackendplugin.azurewebsites.net/fetchCustomerProfile?outlookEmailId=${encodedEmailId}`
      );
      const result = await response.json();
      return result.customerProfile;
    } catch (error) {
      console.error("Error fetching customer profile from backend:", error);
      return "Error fetching customer profile.";
    }
  };

  const fetchObjectNameFromCosmosDB = async (outlookEmailId: string) => {
    try {
      const encodedEmailId = encodeURIComponent(outlookEmailId);
      const response = await fetch(
        `https://cosmosdbbackendplugin.azurewebsites.net/fetchName?outlookEmailId=${encodedEmailId}`
      );
      const result = await response.json();
      return result.objectname;
    } catch (error) {
      console.error("Error fetching objectname from CosmosDB:", error);
      return "Error fetching objectname.";
    }
  };

  // Function to check if the "akzeptiert" folder exists, and create it if not
  const ensureAkzeptiertFolderExists = async (): Promise<string | null> => {
    try {
      // Check if the folder exists
      const response = await axios.get(
        "https://graph.microsoft.com/v1.0/me/mailFolders",
        {
          headers: {
            Authorization: `Bearer ${accessToken}`,
          },
        }
      );

      const folders = response.data.value;
      let folder = folders.find((f: any) => f.displayName === "akzeptiert");

      if (folder) {
        // Folder exists, return its ID
        return folder.id;
      } else {
        // Folder doesn't exist, create it
        const createFolderResponse = await axios.post(
          "https://graph.microsoft.com/v1.0/me/mailFolders",
          {
            displayName: "akzeptiert",
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
      console.error("Error ensuring 'akzeptiert' folder exists:", error);
      return null;
    }
  };

  // Function to create a draft reply and move it to the "akzeptiert" folder
  const createDraftReplyAndMove = async () => {
    try {
      // Fetch all emails with the same folder name from Cosmos DB
      const folderName = await fetchFolderNameFromBackend(emailId);
  
      if (!folderName) {
        console.error("Could not obtain folder name from Cosmos DB.");
        return;
      }
      let emails = await fetchEmailsByFolderName(folderName);
  
      if (!emails || emails.length === 0) {
        console.log(`No emails found with folder name: ${folderName}`);
        return;
      }
  
      // Parse requestInput to get the number of accepted emails
      const numberOfAcceptedEmails = parseInt(requestInput, 10);
  
      if (isNaN(numberOfAcceptedEmails) || numberOfAcceptedEmails < 0) {
        console.error("Invalid number in requestInput");
        return;
      }
  
      // Sort emails by rating in descending order
      emails.sort((a, b) => b.rating - a.rating);
  
      // Split the emails into accepted and rejected arrays
      const acceptedEmails = emails.filter(email => email.rating > 7).slice(0, numberOfAcceptedEmails);
      const rejectedEmails = emails.filter(email => email.rating <= 7).concat(emails.slice(numberOfAcceptedEmails));
  
      // Ensure 'akzeptiert' and 'abgelehnt' folders exist and get their IDs
      const acceptedFolderId = await ensureFolderExists("akzeptiert"+folderName);
      const rejectedFolderId = await ensureFolderExists("abgelehnt"+folderName);
  
      // For accepted emails, create drafts with confirmationTemplate
      for (const email of acceptedEmails) {
        await createDraftReplyForEmail(email.outlookEmailId, acceptedFolderId, confirmationTemplate);
      }
  
      // For rejected emails, create drafts with rejectionTemplate
      for (const email of rejectedEmails) {
        await createDraftReplyForEmail(email.outlookEmailId, rejectedFolderId, rejectionTemplate);
      }
  
      console.log("Draft replies created and moved to the appropriate folders.");
  
      // Check the toggle before deleting the original emails from the inbox
      if (deleteEmailsToggle) {
        for (const email of emails) {
          await deleteEmail(email.outlookEmailId);
        }
      }
  
      // Switch to Frame3
      switchToFrame3(requestInput);
    } catch (error) {
      console.error("Error creating draft replies:", error);
    }
  };
  

  const fetchFolderNameFromBackend = async (outlookEmailId: string) => {
    try {
      const encodedEmailId = encodeURIComponent(outlookEmailId);
      const response = await fetch(
        `https://cosmosdbbackendplugin.azurewebsites.net/fetchFolderName?outlookEmailId=${encodedEmailId}`
      );
      const result = await response.json();
      return result.folderName;
    } catch (error) {
      console.error("Error fetching folder name from backend:", error);
      return null;
    }
  };
  
  const ensureFolderExists = async (folderName: string): Promise<string | null> => {
    try {
      // Check if the folder exists
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
        // Folder exists, return its ID
        return folder.id;
      } else {
        // Folder doesn't exist, create it
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
  const fetchEmailsByFolderName = async (folderName: string) => {
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

  const createDraftReplyForEmail = async (emailId: string, folderId: string, template: string) => {
    try {
      // Create the draft reply with the provided template
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
  
      // Move the draft to the specified folder
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
  
  const toggleDeleteEmails = () => {
    const newToggle = !deleteEmailsToggle;
    setDeleteEmailsToggle(newToggle);
    localStorage.setItem("deleteEmailsToggle", newToggle.toString());
    console.log(`Delete emails toggle set to: ${newToggle}`);
  };
  
  useEffect(() => {
    const fetchEmailContent = async () => {
      if (Office.context.mailbox.item && Office.context.mailbox.item.itemId) {
        // Get the REST ID of the current email
        const restId = Office.context.mailbox.convertToRestId(
          Office.context.mailbox.item.itemId,
          Office.MailboxEnums.RestVersion.v2_0
        );

        // Fetch customer profile and property name
        const customerProfile = await fetchCustomerProfileFromBackend(restId);
        setCustomerProfile(customerProfile);
        const objectname = await fetchObjectNameFromCosmosDB(restId);
        setPropertyName(objectname);

        // Fetch folder name and number of emails with the same folder
        const folderName = await fetchFolderNameFromBackend(restId);
        if (folderName) {
          let emails = await fetchEmailsByFolderName(folderName);
          console.log("Fetched emails:", emails);

          // Sort emails by rating in descending order
          emails = emails.sort((a, b) => b.rating - a.rating);

          setEmails(emails); // Store the sorted emails in state
          const numberOfEmails = emails.length;

          // Parse requestInput to get the number of accepted emails
          const numberOfAcceptedEmails = parseInt(requestInput, 10) || 0;
          setRequestsInfo(`${numberOfAcceptedEmails} der ${numberOfEmails} Anfragen treffen auf die Profilbeschreibung zu`);
        } else {
          setRequestsInfo(`0 der 0 Anfragen treffen auf die Profilbeschreibung zu`);
        }
      }
    };

    fetchEmailContent();

    const itemChangedHandler = () => {
      fetchEmailContent();
    };

    Office.context.mailbox.addHandlerAsync(
      Office.EventType.ItemChanged,
      itemChangedHandler
    );

    return () => {
      Office.context.mailbox.removeHandlerAsync(
        Office.EventType.ItemChanged,
        itemChangedHandler
      );
    };
  }, [emailId, requestInput]);


  useEffect(() => {
    let intervalId: NodeJS.Timeout | null = null;
  
    const fetchAndSetProfile = async () => {
      const profile = await fetchCustomerProfileFromBackend(emailId);
      const name = await fetchObjectNameFromCosmosDB(emailId);
  
      if (profile && name) {
        setCustomerProfile(profile);
        setPropertyName(name);
      }
    };
  
    const startInitialPolling = async () => {
      for (let i = 0; i < 5; i++) {
        await fetchAndSetProfile();
  
        // Stop the loop if customerProfile is updated
        if (customerProfile) {
          return;
        }
  
        // Wait for 1 second before the next attempt
        await new Promise((resolve) => setTimeout(resolve, 1000));
      }
    };
  
    const startRegularPolling = () => {
      // Poll every 2 seconds
      intervalId = setInterval(fetchAndSetProfile, 2000);
    };
  
    if (!customerProfile) {
      // Initial polling for up to 5 seconds
      startInitialPolling().then(() => {
        if (!intervalId) {
          startRegularPolling();
        }
      });
    } else {
      startRegularPolling();
    }
  
    return () => {
      // Clean up on component unmount
      if (intervalId) {
        clearInterval(intervalId);
      }
    };
  }, [emailId, customerProfile]);
  
  useEffect(() => {
    const validateForm = () => {
      const isConfirmationTemplateValid = confirmationTemplate.trim() !== "";
      const isRejectionTemplateValid = rejectionTemplate.trim() !== "";

      setShowConfirmationTemplateError(!isConfirmationTemplateValid);
      setShowRejectionTemplateError(!isRejectionTemplateValid);

      setIsFormValid(
        isConfirmationTemplateValid &&
        isRejectionTemplateValid
      );
    };

    validateForm();
  }, [confirmationTemplate, rejectionTemplate]);

  // Function to handle MarkdownCard click
  const handleCardClick = (email) => {
    console.log("Clicked email:", email);
    openEmailItem(email.outlookEmailId);
  };
  
  const openEmailItem = (itemId: string) => {
    try {
      // Convert the item ID to the required format
      const restId = Office.context.mailbox.convertToRestId(
        itemId,
        Office.MailboxEnums.RestVersion.v1_0
      );
  
      // Use displayMessageForm to open the email
      Office.context.mailbox.displayMessageForm(restId);
    } catch (error) {
      console.error("Error opening email item:", error);
    }
  };
  

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
