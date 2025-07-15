async function uploadEmailAndAttachments() {
  await Office.onReady();
  const item = Office.context.mailbox.item;

  const subject = item.subject || "bericht";
  const body = await getBodyAsync();
  const attachments = item.attachments || [];

  // Vereenvoudigd: e-mail opslaan als tekstbestand
  const emailContent = `Subject: ${subject}\n\n${body}`;

  await uploadToSharePoint(`${subject}.txt`, emailContent);

  for (const att of attachments) {
    const file = await getAttachment(att.id);
    await uploadToSharePoint(att.name, file, true);
  }

  alert("E-mail en bijlagen opgeslagen in SharePoint");
}

function getBodyAsync() {
  return new Promise((resolve) => {
    Office.context.mailbox.item.body.getAsync("text", (result) => {
      resolve(result.value);
    });
  });
}

async function getAttachment(id) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, (result) => {
      if (result.status === "succeeded") {
        const token = result.value;
        const itemId = Office.context.mailbox.item.itemId;
        const url = `${Office.context.mailbox.restUrl}/v2.0/me/messages/${itemId}/attachments/${id}/$value`;

        fetch(url, {
          headers: { Authorization: `Bearer ${token}` }
        })
          .then((res) => res.blob())
          .then(resolve)
          .catch(reject);
      } else {
        reject(result.error);
      }
    });
  });
}

async function uploadToSharePoint(filename, content, isBlob = false) {
  const accessToken = await getGraphAccessToken();
  const siteId = "me"; // Vervang met juiste waarde
  const driveId = "me"; // Vervang met juiste waarde
  const folderPath = "Documents/add-in";

  //const uploadUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/root:/${folderPath}/${filename}:/content`;
  const uploadUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${folderPath}/${filename}:/content`;
  
  await fetch(uploadUrl, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      ...(isBlob ? {} : { "Content-Type": "text/plain" })
    },
    body: content
  });
}

async function getGraphAccessToken() {
  return new Promise((resolve, reject) => {
    Office.context.auth.getAccessTokenAsync({ allowSignInPrompt: true }, (result) => {
      if (result.status === "succeeded") {
        resolve(result.value);
      } else {
        reject("Geen toegangstoken verkregen");
      }
    });
  });
}
