export const oneDriveConfig = {
  enabled: false,
  clientId: "YOUR-ENTRA-APP-CLIENT-ID",
  authority: "https://login.microsoftonline.com/organizations",
  redirectUri: "https://christian-ori-ai.github.io/IRDR/",
  postLogoutRedirectUri: "https://christian-ori-ai.github.io/IRDR/",
  graphScopes: ["Files.ReadWrite"],
  uploadPath: "IRDR/Results",
  autoUploadOnFinish: true,
};
