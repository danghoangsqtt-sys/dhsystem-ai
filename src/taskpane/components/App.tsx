import * as React from "react";
import { 
  makeStyles, 
  shorthands, 
  Button, 
  Input, 
  Label, 
  Tab, 
  TabList, 
  Text,
  Card,
  CardHeader,
  Textarea,
  Spinner,
  Toaster,
  useId,
  useToastController,
  Toast,
  ToastTitle,
  ToastBody,
  ToastIntent
} from "@fluentui/react-components";
import { 
  Flash24Regular, 
  ShieldCheckmark24Regular, 
  EyeRegular, 
  EyeOffRegular,
  Translate24Regular,
  Code24Regular,
  Diagram24Regular,
  Image24Regular,
  CheckmarkCircle24Regular,
  Info24Regular,
  ErrorCircle24Regular
} from "@fluentui/react-icons";
import { translateText, analyzeCode, getMermaidCode, callGemini } from "../gemini-api";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    height: "100vh",
    backgroundColor: "#f5f5f5",
    ...shorthands.padding("0px"),
    boxSizing: "border-box",
  },
  header: {
    backgroundColor: "#ffffff",
    ...shorthands.padding("12px", "16px"),
    borderBottom: "1px solid #e0e0e0",
    display: "flex",
    flexDirection: "column",
    gap: "8px",
  },
  titleArea: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
  },
  logoGroup: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
  },
  logo: {
    backgroundColor: "#2b579a",
    color: "white",
    ...shorthands.padding("4px"),
    ...shorthands.borderRadius("6px"),
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
  },
  premiumBadge: {
    backgroundColor: "#ebf3fc",
    color: "#2b579a",
    ...shorthands.padding("2px", "8px"),
    ...shorthands.borderRadius("12px"),
    fontSize: "10px",
    fontWeight: "bold",
    textTransform: "uppercase",
  },
  apiSection: {
    display: "flex",
    flexDirection: "column",
    gap: "4px",
    marginTop: "8px",
  },
  apiInputGroup: {
    display: "flex",
    gap: "8px",
  },
  nav: {
    backgroundColor: "#ffffff",
    borderBottom: "1px solid #f0f0f0",
  },
  main: {
    flex: 1,
    overflowY: "auto",
    ...shorthands.padding("16px"),
    display: "flex",
    flexDirection: "column",
    gap: "20px",
  },
  card: {
    ...shorthands.padding("16px"),
    display: "flex",
    flexDirection: "column",
    gap: "12px",
  },
  footer: {
    ...shorthands.padding("8px", "16px"),
    backgroundColor: "#ffffff",
    borderTop: "1px solid #e0e0e0",
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
  },
  statusIndicator: {
    display: "flex",
    alignItems: "center",
    gap: "6px",
    fontSize: "10px",
    color: "#616161",
    fontWeight: "bold",
    textTransform: "uppercase",
  },
  dot: {
    width: "6px",
    height: "6px",
    borderRadius: "50%",
  },
  textarea: {
    width: "100%",
    minHeight: "120px",
  },
  buttonGroup: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
  },
  resultArea: {
    marginTop: "12px",
    padding: "12px",
    backgroundColor: "#f9f9f9",
    border: "1px solid #e0e0e0",
    borderRadius: "8px",
    fontSize: "12px",
    whiteSpace: "pre-wrap",
  },
  loader: {
    display: "flex",
    justifyContent: "center",
    ...shorthands.padding("20px"),
  },
  apiLabel: {
    color: "#888", 
    fontSize: "10px", 
    letterSpacing: "1px",
  },
  apiInput: {
    flexGrow: 1,
  },
  footerText: {
    color: "#ccc",
  }
});

const App: React.FC = () => {
  const styles = useStyles();
  const toasterId = useId("toaster");
  const { dispatchToast } = useToastController(toasterId);

  const [apiKey, setApiKey] = React.useState<string>(localStorage.getItem("gemini_api_key") || "");
  const [showKey, setShowKey] = React.useState<boolean>(false);
  const [activeTab, setActiveTab] = React.useState<string>("translate");
  const [loading, setLoading] = React.useState<boolean>(false);
  const [status, setStatus] = React.useState<{message: string, intent: ToastIntent}>({
    message: "System Ready",
    intent: "success"
  });

  const notify = (message: string, intent: ToastIntent = "info") => {
    dispatchToast(
      <Toast>
        <ToastTitle>{intent === "success" ? "Success" : intent === "error" ? "Error" : "Info"}</ToastTitle>
        <ToastBody>{message}</ToastBody>
      </Toast>,
      { intent }
    );
    setStatus({ message, intent });
  };

  const handleSaveKey = () => {
    localStorage.setItem("gemini_api_key", apiKey);
    notify("API Key saved successfully", "success");
  };

  const wrapAction = async (action: () => Promise<void>) => {
    if (!apiKey) {
      notify("Please configure Gemini API Key first", "error");
      return;
    }
    setLoading(true);
    try {
      await action();
    } catch (error: any) {
      notify(error.message || "An error occurred", "error");
    } finally {
      setLoading(false);
    }
  };

  // Views Logic
  const [translateInput, setTranslateInput] = React.useState("");
  const [codeHighlight, setCodeHighlight] = React.useState("");
  const [diagramPrompt, setDiagramPrompt] = React.useState("");

  const onTranslate = () => wrapAction(async () => {
    const result = await translateText(translateInput, apiKey);
    await Office.context.document.setSelectedDataAsync(result);
    notify("Translation inserted", "success");
  });

  const onAnalyze = () => wrapAction(async () => {
    const result = await analyzeCode(codeHighlight, apiKey);
    notify("Analysis complete. See console for now.", "success");
    console.log(result);
  });

  return (
    <div className={styles.container}>
      <Toaster toasterId={toasterId} />
      
      {/* Step 1: Luxurious Header */}
      <header className={styles.header}>
        <div className={styles.titleArea}>
          <div className={styles.logoGroup}>
            <div className={styles.logo}>
              <Flash24Regular />
            </div>
            <Text weight="bold" size={400}>DhSystem AI</Text>
          </div>
          <div className={styles.premiumBadge}>Premium</div>
        </div>

        <div className={styles.apiSection}>
          <Label size="small" weight="semibold" className={styles.apiLabel}>API CONFIGURATION</Label>
          <div className={styles.apiInputGroup}>
            <Input 
              type={showKey ? "text" : "password"} 
              value={apiKey} 
              onChange={(e, data) => setApiKey(data.value)}
              placeholder="Gemini API Key..."
              contentAfter={
                <Button 
                  appearance="transparent" 
                  icon={showKey ? <EyeOffRegular /> : <EyeRegular />} 
                  onClick={() => setShowKey(!showKey)}
                />
              }
              className={styles.apiInput}
            />
            <Button appearance="primary" onClick={handleSaveKey}>Save</Button>
          </div>
        </div>
      </header>

      {/* Navigation */}
      <nav className={styles.nav}>
        <TabList selectedValue={activeTab} onTabSelect={(e, data) => setActiveTab(data.value as string)}>
          <Tab id="translate" value="translate" icon={<Translate24Regular />}>Translate</Tab>
          <Tab id="code" value="code" icon={<Code24Regular />}>Code</Tab>
          <Tab id="diagram" value="diagram" icon={<Diagram24Regular />}>Draw</Tab>
          <Tab id="art" value="art" icon={<Image24Regular />}>Art</Tab>
        </TabList>
      </nav>

      {/* Main Content */}
      <main className={styles.main}>
        {activeTab === "translate" && (
          <Card className={styles.card}>
            <CardHeader 
              image={<Translate24Regular />} 
              header={<Text weight="semibold">Smart Translation</Text>}
              description="Translate technical/IT content"
            />
            <Textarea 
              className={styles.textarea} 
              value={translateInput}
              onChange={(e, data) => setTranslateInput(data.value)}
              placeholder="Enter text or select in Word..."
            />
            <div className={styles.buttonGroup}>
              <Button appearance="primary" icon={<Flash24Regular />} disabled={loading} onClick={onTranslate}>
                Expert Trans
              </Button>
            </div>
          </Card>
        )}

        {activeTab === "code" && (
          <Card className={styles.card}>
            <CardHeader 
              image={<Code24Regular />} 
              header={<Text weight="semibold">Code Intelligence</Text>}
            />
            <Textarea 
              className={styles.textarea} 
              value={codeHighlight}
              onChange={(e, data) => setCodeHighlight(data.value)}
              placeholder="// Paste snippet..."
              style={{fontFamily: "monospace"}}
            />
            <Button appearance="primary" onClick={onAnalyze} disabled={loading}>Analyze Bugs</Button>
          </Card>
        )}

        {loading && (
          <div className={styles.loader}>
            <Spinner label="AI is thinking..." />
          </div>
        )}
      </main>

      {/* Footer Status */}
      <footer className={styles.footer}>
        <div className={styles.statusIndicator}>
          <div className={styles.dot} style={{backgroundColor: status.intent === "error" ? "#ff0000" : "#10b981"}}></div>
          <Text size={100}>{status.message}</Text>
        </div>
        <Text size={100} className={styles.footerText}>V2.5.4</Text>
      </footer>
    </div>
  );
};

export default App;
