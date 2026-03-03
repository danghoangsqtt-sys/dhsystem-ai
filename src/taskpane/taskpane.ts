/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "block";

    // Initialize Tab Switching
    initTabs();

    // Initialize Button Handlers
    initButtonHandlers();
  }
});

function initTabs() {
  const tabButtons = document.querySelectorAll(".tab-btn");
  const tabContents = document.querySelectorAll(".tab-content");

  tabButtons.forEach((btn) => {
    btn.addEventListener("click", () => {
      const targetTabId = btn.getAttribute("data-tab");

      // Update active button
      tabButtons.forEach((b) => b.classList.remove("active"));
      btn.classList.add("active");

      // Update active content
      tabContents.forEach((content) => {
        content.classList.remove("active");
        if (content.id === targetTabId) {
          content.classList.add("active");
        }
      });
    });
  });
}

function initButtonHandlers() {
  // Tab 1: Translation
  document.getElementById("btn-translate-pro").onclick = () =>
    handleAction("Dịch thuật chuyên ngành");
  document.getElementById("btn-fix-grammar").onclick = () => handleAction("Sửa lỗi chính tả");

  // Tab 2: Code
  document.getElementById("btn-explain-code").onclick = () => handleAction("Giải thích Code");
  document.getElementById("btn-analyze-bugs").onclick = () => handleAction("Phân tích & Tìm lỗi");

  // Tab 3: Diagram
  document.getElementById("btn-generate-diagram").onclick = () => handleAction("Tạo & Chèn Sơ đồ");

  // Tab 4: Image
  document.getElementById("btn-search-image").onclick = () => handleAction("Tìm ảnh Online");
  document.getElementById("btn-ai-image").onclick = () => handleAction("AI Generate Ảnh");
}

async function handleAction(actionName: string) {
  console.log(`Action triggered: ${actionName}`);

  // Example of interacting with Word
  return Word.run(async (context) => {
    const range = context.document.getSelection();
    range.load("text");
    await context.sync();

    console.log(`Selected text for ${actionName}: ${range.text}`);

    // For now, we just show an alert or placeholder logic
    // In a real app, this would call an AI API
    if (!range.text) {
      // If nothing is selected, maybe insert a placeholder
      const paragraph = context.document.body.insertParagraph(
        `[Đang xử lý ${actionName}...]`,
        Word.InsertLocation.end
      );
      paragraph.font.bold = true;
    } else {
      // Process selected text
      const paragraph = range.insertParagraph(
        `[Kết quả ${actionName} cho: "${range.text.substring(0, 20)}..."]`,
        Word.InsertLocation.after
      );
      paragraph.font.italic = true;
      paragraph.font.color = "blue";
    }

    await context.sync();
  });
}
