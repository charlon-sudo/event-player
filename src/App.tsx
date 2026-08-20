import "./App.css";
import { useMsal } from "@azure/msal-react";

function App() {
  const { instance, accounts } = useMsal();

  const signIn = async () => {
    await instance.loginRedirect({
      scopes: ["User.Read"],
    });
  };

  const signOut = async () => {
    await instance.logoutRedirect();
  };
  const loadMedia = async () => {
  try {
    const tokenResponse =
      await instance.acquireTokenSilent({
        scopes: [
          "User.Read",
          "Files.Read.All",
          "Sites.Read.All",
        ],
        account: accounts[0],
      });

    const response = await fetch(
      "https://graph.microsoft.com/v1.0/sites/grandhotelwarrandyte.sharepoint.com:/sites/GHW",
      {
        headers: {
          Authorization: `Bearer ${tokenResponse.accessToken}`,
        },
      }
    );

    const data = await response.json();
console.log(data);
    console.log("SITE DATA", data);

const driveResponse = await fetch(
  `https://graph.microsoft.com/v1.0/sites/${data.id}/drive`,
  {
    headers: {
      Authorization: `Bearer ${tokenResponse.accessToken}`,
    },
  }
);

const driveData = await driveResponse.json();

console.log("DRIVE DATA", driveData);

console.log("DRIVE DATA", driveData);

const folderResponse = await fetch(
  `https://graph.microsoft.com/v1.0/drives/${driveData.id}/root:/Functions/Functions%20AV%20Uploads/Caledonia%20Room/Event%201:/children`,
  {
    headers: {
      Authorization: `Bearer ${tokenResponse.accessToken}`,
    },
  }
);

const folderData = await folderResponse.json();

console.log("FOLDER DATA", folderData);

const fileNames = folderData.value
  .map((file: any) => file.name)
  .join("\n");

alert(fileNames);
  } catch (error) {
    console.error(error);
    alert("Failed");
  }
};

  return (
    <div className="app">
      <div className="container">
        <h1>Event Player</h1>

        {accounts.length > 0 ? (
          <>
            <h2>Welcome</h2>

            <p>{accounts[0].username}</p>

           <div className="buttonRow">
  <button
  className="signInButton"
  onClick={loadMedia}
>
  Load SharePoint Media
</button>

  <button
    className="secondaryButton"
    onClick={signOut}
  >
    Sign Out
  </button>
</div>
          </>
        ) : (
          <button
            className="signInButton"
            onClick={signIn}
          >
            Sign in with Microsoft
          </button>
        )}
      </div>
    </div>
  );
}

export default App;