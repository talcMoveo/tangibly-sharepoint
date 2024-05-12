import express from "express";
import fetch from "node-fetch";
import moment from "moment";
import fs from "fs";

const app = express();
app.use(express.json());
const PORT = 3000;

const siteUrl = "https://talcos.sharepoint.com/sites/Taltest";
const listId = "8868b46c-fb06-4290-a78f-be063c71e4e4";
const accessToken =
  "eyJ0eXAiOiJKV1QiLCJub25jZSI6ImVFdGM5WTFXNkN3TzFpOEhtNWJGVDNpOC1QcHA2eVpkWTBzcWZvNmJWb28iLCJhbGciOiJSUzI1NiIsIng1dCI6InEtMjNmYWxldlpoaEQzaG05Q1Fia1A1TVF5VSIsImtpZCI6InEtMjNmYWxldlpoaEQzaG05Q1Fia1A1TVF5VSJ9.eyJhdWQiOiJodHRwczovL3RhbGNvcy5zaGFyZXBvaW50LmNvbSIsImlzcyI6Imh0dHBzOi8vc3RzLndpbmRvd3MubmV0LzY1YzBlZDI0LWRhZDktNGUzMy1hYjViLTliMzcxYmVmYjFiZS8iLCJpYXQiOjE3MTMzNjQ2OTgsIm5iZiI6MTcxMzM2NDY5OCwiZXhwIjoxNzEzMzY5MzIxLCJhY3IiOiIxIiwiYWlvIjoiQVRRQXkvOFdBQUFBN1hWQVEyTzBMam81MUZaQWZPR1dhNDQvTFFTbVFjSGcrMzh1UFV4UWNWSW9KMGhWTDF2Sy83eVJOb0ZQMGpqUCIsImFtciI6WyJwd2QiXSwiYXBwX2Rpc3BsYXluYW1lIjoic2hhcmVwb2ludCIsImFwcGlkIjoiN2YzZDk3Y2UtMTJlNi00ZmFiLThkN2MtODRhNGRkN2Y2NWI5IiwiYXBwaWRhY3IiOiIxIiwiZmFtaWx5X25hbWUiOiLXm9eU158iLCJnaXZlbl9uYW1lIjoi15jXnCIsImlkdHlwIjoidXNlciIsImlwYWRkciI6IjkzLjE1Ny44Ni4xMzQiLCJuYW1lIjoi15jXnCDXm9eU158iLCJvaWQiOiIyYjU0Njc2Yy04NzViLTQ1YzEtOGVmOS1iYzQyZTUxYTAxZGUiLCJwdWlkIjoiMTAwMzIwMDM2RTA0RUY4MSIsInJoIjoiMC5BYThBSk8zQVpkbmFNMDZyVzVzM0ctLXh2Z01BQUFBQUFQRVB6Z0FBQUFBQUFBQWRBVFEuIiwic2NwIjoiQWxsU2l0ZXMuRnVsbENvbnRyb2wgQWxsU2l0ZXMuTWFuYWdlIEFsbFNpdGVzLlJlYWQgQWxsU2l0ZXMuV3JpdGUgRW50ZXJwcmlzZVJlc291cmNlLlJlYWQgRW50ZXJwcmlzZVJlc291cmNlLldyaXRlIE15RmlsZXMuUmVhZCBNeUZpbGVzLldyaXRlIFByb2plY3QuUmVhZCBQcm9qZWN0LldyaXRlIFByb2plY3RXZWJBcHAuRnVsbENvbnRyb2wgUHJvamVjdFdlYkFwcFJlcG9ydGluZy5SZWFkIFNpdGVzLlNlYXJjaC5BbGwgVGFza1N0YXR1cy5TdWJtaXQgVGVybVN0b3JlLlJlYWQuQWxsIFRlcm1TdG9yZS5SZWFkV3JpdGUuQWxsIFVzZXIuUmVhZCBVc2VyLlJlYWQuQWxsIFVzZXIuUmVhZFdyaXRlLkFsbCIsInNpZCI6IjMxZWYwYjhjLTEyZDEtNDVlMC04Y2NiLTI0MzE5MjJkZmRjMSIsInNpZ25pbl9zdGF0ZSI6WyJrbXNpIl0sInN1YiI6IktqaWZpYkhENGxtdWk3Sm9jZjhJdHZrWTBNVllJWEprRjF1b012djN4TzgiLCJ0aWQiOiI2NWMwZWQyNC1kYWQ5LTRlMzMtYWI1Yi05YjM3MWJlZmIxYmUiLCJ1bmlxdWVfbmFtZSI6InRhbGNvQHRhbGNvcy5vbm1pY3Jvc29mdC5jb20iLCJ1cG4iOiJ0YWxjb0B0YWxjb3Mub25taWNyb3NvZnQuY29tIiwidXRpIjoiS1ZtSzhidXBYMFdfekpzQmhfb25BQSIsInZlciI6IjEuMCIsIndpZHMiOlsiNjJlOTAzOTQtNjlmNS00MjM3LTkxOTAtMDEyMTc3MTQ1ZTEwIiwiYjc5ZmJmNGQtM2VmOS00Njg5LTgxNDMtNzZiMTk0ZTg1NTA5Il19.plfldR-7T_EGYfHs-Rx4cltNI_6l4SJesKRyro14A7jDlVUbtPf4el6eJiRLdB5h1PTcce6uAEvn9FQk6INdFSE2TIsgXpYv2SfOGNtpWil7w5zbhqHqqClybl965XP8r4tC6zmWpEulSWIhms6qjJvXL0rdKN1EQa9m8hQUm2vRaNuE98dd__DTMsIJn1o0Wc1sjEvIDQFZd215pvVfyAX5GHg5DsN7Yya4IIJDse9MXr2AuOu968oTDnCgnCreMtPDBAu2pT5KEOIGEMROQnBOHCzO9lgou0VXB6Hew5Gv0Bo1YVJ8IK90yuIhEShMPVgPeA8c4PDrq3jbakOAAg";

app.get("/", (req, res) => {
  console.log("application running");
  // try {
  //   initializeShareDocument();
  //   res.send("Share operation successful, check logs for details.");
  // } catch (error) {
  //   console.error("Failed to share document due to an error:", error);
  //   res
  //     .status(500)
  //     .send("Failed to share document. Check server logs for more details.");
  // }
  res.send("Working.");
});

app.listen(PORT, () => {
  console.log(`Express server running at http://localhost:${PORT}/`);
});

app.get("/oauth-redirect", (req, res) => {
  res.send("OAuth response received.");
});

// change timestamp to SP ticks
function toTicks(momentDate) {
  const ticksPerMillisecond = 10000;
  const ticksSinceEpoch = momentDate.valueOf() * ticksPerMillisecond;
  const ticksSince1AD = 621355968000000000;
  return ticksSinceEpoch + ticksSince1AD;
}

// enpoint - recieve webhook
app.post("/api/webhook", async (req, res) => {
  console.log("Webhook received:", req.body);

  if (req.query.validationtoken) {
    return res.send(req.query.validationtoken.toString());
  }

  try {
    const oneHourAgoToken = `1;3;${listId};${toTicks(
      moment().subtract(1, "hours")
    )};-1`;

    const changeQuery = {
      query: {
        Item: true,
        Add: true,
        Update: true,
        DeleteObject: true,
        Restore: true,
        ChangeTokenStart: {
          StringValue: oneHourAgoToken,
        },
      },
    };
    const changesResponse = await fetch(
      `${siteUrl}/_api/web/lists(guid'${listId}')/getChanges`,
      {
        method: "POST",
        headers: {
          Accept: "application/json;odata=nometadata",
          "Content-Type": "application/json",
          Authorization: `Bearer ${accessToken}`,
        },
        body: JSON.stringify(changeQuery),
      }
    );

    const changesData = await changesResponse.json();
    manageChanges(changesData);
  } catch (error) {
    console.error("Error:", error);
  }

  res.status(200).send();
});

// get changes from LAST webhook changed file
function manageChanges(changes) {
  console.log("Changes received:", changes.value);
  const newValues = changes.value;
  newValues.forEach((element, index) => {
    if (index === newValues.length - 1) {
      console.log("Change:", element.UniqueId);
      // shareDocument(
      //   "/sites/Taltest/_layouts/15/Doc.aspx?sourcedoc=%7BEFE0E133-DCA6-4782-AA5E-7BA329303DC9%7D&file=newDocTest.docx",
      //   "talorcoh@gmail.com",
      //   "1",
      //   "hi"
      // );
      // getFileContents(element.UniqueId)
      //   .then((text) => {
      //     console.log("File contents as text:", text);
      //   })
      //   .catch((error) => {
      //     console.error("Error fetching the file:", error);
      //   });
    }
  });
}

// translate file form binary to text
async function getFileContents(fileId) {
  const response = await fetch(
    `${siteUrl}/_api/web/GetFileById('${fileId}')/$value`,
    {
      method: "GET",
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    }
  );

  if (!response.ok) {
    throw new Error("Network response was not ok");
  }
  const buffer = await response.buffer();
  const filePath = `downloadedFile-${fileId}`;
  fs.writeFileSync(filePath, buffer);
  return `File saved to ${filePath}`;
}

// SHARE DOCUMENT
async function initializeShareDocument() {
  const shareResponse = await shareDocument(
    "https://talcos.sharepoint.com/sites/Taltest/Shared%20Documents/add.docx",
    "danielbar@moveo.co.il",
    "1",
    "Please review this document."
  );
  console.log("Document sharing successful:", shareResponse);
}

// Function to handle document sharing
async function shareDocument(resourceAddress, userId, role, customMessage) {
  const requestDigest = await getRequestDigest();
  const body = {
    resourceAddress: resourceAddress,
    userRoleAssignments: [
      {
        __metadata: { type: "SP.Sharing.UserRoleAssignment" },
        Role: role,
        UserId: userId,
      },
    ],
    validateExistingPermissions: false,
    additiveMode: true,
    sendServerManagedNotification: false,
    customMessage: customMessage,
    includeAnonymousLinksInNotification: false,
  };

  const url = `${siteUrl}/_api/SP.Sharing.DocumentSharingManager.UpdateDocumentSharingInfo`;
  try {
    const response = await fetch(url, {
      method: "POST",
      headers: {
        Accept: "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        Authorization: `Bearer ${accessToken}`,
        "X-RequestDigest": requestDigest,
      },
      body: JSON.stringify(body),
    });

    if (!response.ok) {
      const text = await response.text(); // Get response body for detailed error
      throw new Error(
        `SharePoint API call failed with status ${response.status}: ${text}`
      );
    }
    return await response.json();
  } catch (error) {
    console.error("Error sharing document:", error);
    throw error;
  }
}

// Function to get the SharePoint request digest
async function getRequestDigest() {
  const url = `${siteUrl}/_api/contextinfo`;
  try {
    const response = await fetch(url, {
      method: "POST",
      headers: {
        Accept: "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        Authorization: `Bearer ${accessToken}`,
      },
    });
    const data = await response.json();
    return data.d.GetContextWebInformation.FormDigestValue;
  } catch (error) {
    console.error("Error fetching request digest:", error);
    throw error;
  }
}
