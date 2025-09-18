function insertLogo(event) {
  // Fixed image (use GitHub Pages logo URL once uploaded)
  const imageUrl = "https://yourusername.github.io/practice-outlook-addin/logo.png";

  // Ask for hyperlink
  const hyperlink = prompt("Enter the hyperlink (e.g., https://www.practicearch.com):");
  if (!hyperlink) {
    if (event) event.completed();
    return;
  }

  // Insert the image wrapped in a link
  Office.context.mailbox.item.body.setSelectedDataAsync(
    `<a href="${hyperlink}" target="_blank"><img src="${imageUrl}" style="max-width:200px;"/></a>`,
    { coercionType: Office.CoercionType.Html },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Failed to insert logo: " + asyncResult.error.message);
      }
      if (event) event.completed();
    }
  );
}
