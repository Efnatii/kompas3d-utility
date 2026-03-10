import { createExecutorShell } from "./executor-shell.js";
import { createXlsxToKompasTblModule } from "./modules/xlsx-to-kompas-tbl.js";

function createPlaceholderModule({ id, title, tabLabel, tabDetail, message }) {
  return {
    id,
    title,
    subtitle: "Каркас вкладки уже готов, логика будет добавляться как отдельный модуль без переписывания shell.",
    tabLabel,
    tabDetail,
    getRuntimeContribution() {
      return {
        commands: {},
        allowedTypes: [],
        allowedProcesses: [],
      };
    },
    mount(container, context) {
      container.innerHTML = `
        <div class="placeholder">
          <h2 class="module-title">${title}</h2>
          <p>${message}</p>
        </div>
      `;
      context.setModuleBadge("placeholder", false);
      context.setModuleMeta(message);
      return {};
    },
  };
}

const executor = createExecutorShell({
  modules: [
    createXlsxToKompasTblModule(),
    createPlaceholderModule({
      id: "kompas-text-sync",
      title: "KOMPAS Text Sync",
      tabLabel: "TEXT",
      tabDetail: "next",
      message: "Резерв под web-перенос text-sync утилиты.",
    }),
    createPlaceholderModule({
      id: "utility-queue",
      title: "Utility Queue",
      tabLabel: "QUEUE",
      tabDetail: "todo",
      message: "Резерв под следующие утилиты этого репозитория.",
    }),
  ],
});

executor.init().catch((error) => {
  const logOutput = document.getElementById("log-output");
  if (logOutput) {
    logOutput.textContent += `\n[bootstrap] ${String(error.message || error)}`;
  }
});
