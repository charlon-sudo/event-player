import "./App.css";
import { useEffect, useRef, useState } from "react";
import { useMsal } from "@azure/msal-react";

type MediaType = "image" | "video";

interface SharePointFile {
  id: string;
  name: string;
  "@microsoft.graph.downloadUrl"?: string;
}

interface MediaItem {
  id: string;
  name: string;
  url: string;
  type: MediaType;
}

interface SiteResponse {
  id?: string;
  error?: { message?: string };
}

interface DriveResponse {
  id?: string;
  error?: { message?: string };
}

interface FolderResponse {
  value?: SharePointFile[];
  error?: { message?: string };
}

const GRAPH_BASE = "https:" + "//graph.microsoft.com/v1.0";
const SCOPES = ["User.Read", "Files.Read.All", "Sites.Read.All"];

function App() {
  const { instance, accounts } = useMsal();

  const [mediaItems, setMediaItems] = useState<MediaItem[]>([]);
  const [currentIndex, setCurrentIndex] = useState(0);
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState("");
  const [presentationMode, setPresentationMode] = useState(false);
  const [soundEnabled, setSoundEnabled] = useState(false);

  const imageTimerRef = useRef<number | undefined>(undefined);
  const videoRef = useRef<HTMLVideoElement | null>(null);

  const currentItem = mediaItems[currentIndex];

  const signIn = async (): Promise<void> => {
    await instance.loginRedirect({ scopes: SCOPES });
  };

  const signOut = async (): Promise<void> => {
    await instance.logoutRedirect();
  };

  const getMediaType = (fileName: string): MediaType | undefined => {
    const extension = fileName.split(".").pop()?.toLowerCase();
    const imageExtensions = ["jpg", "jpeg", "png", "gif", "webp", "bmp"];
    const videoExtensions = ["mp4", "mov", "m4v", "webm", "ogg", "mkv"];

    if (extension && imageExtensions.includes(extension)) return "image";
    if (extension && videoExtensions.includes(extension)) return "video";
    return undefined;
  };

  const nextSlide = (): void => {
    setCurrentIndex((index) =>
      mediaItems.length > 0 ? (index + 1) % mediaItems.length : 0
    );
  };

  const previousSlide = (): void => {
    setCurrentIndex((index) => {
      if (mediaItems.length === 0) return 0;
      return index === 0 ? mediaItems.length - 1 : index - 1;
    });
  };

  const loadMedia = async (): Promise<void> => {
    const account = accounts[0];
    if (!account) {
      setMessage("Please sign in first.");
      return;
    }

    setLoading(true);
    setMessage("");
    setMediaItems([]);
    setCurrentIndex(0);

    try {
      const tokenResponse = await instance.acquireTokenSilent({
        scopes: SCOPES,
        account,
      });

      const headers = {
        Authorization: `Bearer ${tokenResponse.accessToken}`,
      };

      const siteResponse = await fetch(
        GRAPH_BASE +
          "/sites/grandhotelwarrandyte.sharepoint.com:/sites/GHW",
        { headers }
      );

      if (!siteResponse.ok) {
        throw new Error(
          `Site request failed: ${siteResponse.status} ${siteResponse.statusText}`
        );
      }

      const siteData = (await siteResponse.json()) as SiteResponse;
      if (!siteData.id) {
        throw new Error(siteData.error?.message || "SharePoint site ID not returned.");
      }

      const driveResponse = await fetch(
        GRAPH_BASE + "/sites/" + encodeURIComponent(siteData.id) + "/drive",
        { headers }
      );

      if (!driveResponse.ok) {
        throw new Error(
          `Drive request failed: ${driveResponse.status} ${driveResponse.statusText}`
        );
      }

      const driveData = (await driveResponse.json()) as DriveResponse;
      if (!driveData.id) {
        throw new Error(
          driveData.error?.message || "Document library ID not returned."
        );
      }

      const folderPath =
        "Functions/Functions AV Uploads/Caledonia Room/Event 1";
      const encodedFolderPath = folderPath
        .split("/")
        .map((part) => encodeURIComponent(part))
        .join("/");

      const folderResponse = await fetch(
        GRAPH_BASE +
          "/drives/" +
          encodeURIComponent(driveData.id) +
          "/root:/" +
          encodedFolderPath +
          ":/children",
        { headers }
      );

      if (!folderResponse.ok) {
        throw new Error(
          `Folder request failed: ${folderResponse.status} ${folderResponse.statusText}`
        );
      }

      const folderData = (await folderResponse.json()) as FolderResponse;

      const items = (folderData.value || [])
        .map((file): MediaItem | undefined => {
          const type = getMediaType(file.name);
          const url = file["@microsoft.graph.downloadUrl"];
          if (!type || !url) return undefined;
          return { id: file.id, name: file.name, url, type };
        })
        .filter((item): item is MediaItem => item !== undefined);

      setMediaItems(items);
      setMessage(
        items.length > 0
          ? `Loaded ${items.length} media file(s) from Event 1.`
          : "No supported images or videos were found in Event 1."
      );
    } catch (error) {
      console.error(error);
      setMessage(
        error instanceof Error ? error.message : "Unable to load SharePoint media."
      );
    } finally {
      setLoading(false);
    }
  };

  const startPresentation = (): void => {
  if (mediaItems.length === 0) return;

  setPresentationMode(true);
  setSoundEnabled(false);

  const video = videoRef.current;

  if (video) {
    video.play().catch(() => {});
  }
};
  const exitPresentation = async (): Promise<void> => {
  setPresentationMode(false);
  document.body.style.overflow = "";

  try {
    if (document.fullscreenElement) {
      await document.exitFullscreen();
    }
  } catch (error) {
    console.warn("Exit fullscreen failed", error);
  }
};

  useEffect(() => {
    if (imageTimerRef.current !== undefined) {
      window.clearTimeout(imageTimerRef.current);
    }

    if (
      !presentationMode ||
      !currentItem ||
      currentItem.type !== "image" ||
      mediaItems.length <= 1
    ) {
      return;
    }

    imageTimerRef.current = window.setTimeout(() => {
      setCurrentIndex((index) => (index + 1) % mediaItems.length);
    }, 2000);

    return () => {
      if (imageTimerRef.current !== undefined) {
        window.clearTimeout(imageTimerRef.current);
      }
    };
  }, [currentIndex, currentItem, mediaItems.length, presentationMode]);

  useEffect(() => {
  const video = videoRef.current;

  if (!presentationMode || !video || currentItem?.type !== "video") {
    return;
  }
console.log("PLAYING VIDEO", currentItem.name);
  video.muted = !soundEnabled;

setTimeout(() => {
  video.play().catch((error: unknown) => {
    console.warn(
      "Browser blocked automatic video playback.",
      error
    );
  });
}, 300);
}, [
  currentIndex,
  currentItem,
  presentationMode,
  soundEnabled,
]);

  return (
    <div className={presentationMode ? "app presentationApp" : "app"}>
      {!presentationMode && (
        <div className="container">
          <h1>Event Player</h1>

          {accounts.length > 0 ? (
            <>
              <h2>Welcome</h2>
              <p>{accounts[0].username}</p>

              <div className="buttonRow">
                <button
                  className="primaryButton"
                  type="button"
                  onClick={() => void loadMedia()}
                  disabled={loading}
                >
                  {loading ? "Loading Media..." : "Load SharePoint Media"}
                </button>

                <button
                  className="secondaryButton"
                  type="button"
                  onClick={() => void signOut()}
                >
                  Sign Out
                </button>
              </div>

              {message && <p className="statusMessage">{message}</p>}

              {mediaItems.length > 0 && (
                <>
                  <button
                    className="startButton"
                    type="button"
                    onClick={startPresentation}
                  >
                    Start Presentation
                  </button>

                  <div className="mediaGrid">
                    {mediaItems.map((item) => (
                      <div className="mediaCard" key={item.id}>
                        {item.type === "image" ? (
                          <img src={item.url} alt={item.name} />
                        ) : (
                          <video src={item.url} muted playsInline preload="metadata" />
                        )}
                        <p>{item.name}</p>
                      </div>
                    ))}
                  </div>
                </>
              )}
            </>
          ) : (
            <button
              className="primaryButton"
              type="button"
              onClick={() => void signIn()}
            >
              Sign in with Microsoft
            </button>
          )}
        </div>
      )}

      {presentationMode && currentItem && (
        <div className="presentationStage">
          {currentItem.type === "image" ? (
            <img
              className="presentationMedia"
              src={currentItem.url}
              alt={currentItem.name}
            />
          ) : (
          <video
  key={currentItem.id}
  ref={videoRef}
  className="presentationMedia"
  autoPlay
  muted={!soundEnabled}
  playsInline
  preload="auto"
  
  controls={true}
 onEnded={() => {
  console.log("VIDEO ENDED", currentItem.name);
  nextSlide();
}}
            />
          )}

          <button
            className="previousButton"
            type="button"
            onClick={previousSlide}
            aria-label="Previous media"
          >
            ‹
          </button>

          <button
            className="nextButton"
            type="button"
            onClick={nextSlide}
            aria-label="Next media"
          >
            ›
          </button>

          <div className="presentationControls">
            <button
  type="button"
  onClick={() => {
    const newValue = !soundEnabled;

    setSoundEnabled(newValue);

    if (videoRef.current) {
      videoRef.current.muted = !newValue;
      videoRef.current.currentTime = videoRef.current.currentTime;

      if (newValue) {
        videoRef.current.volume = 1;
      }

      videoRef.current.play().catch((error) => {
        console.warn("Playback retry failed", error);
      });
    }
  }}
>
  {soundEnabled ? "Mute" : "Enable Sound"}
</button>

            <span>
              {currentIndex + 1} / {mediaItems.length}
            </span>

            <button type="button" onClick={exitPresentation}>
              Exit
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

export default App;
