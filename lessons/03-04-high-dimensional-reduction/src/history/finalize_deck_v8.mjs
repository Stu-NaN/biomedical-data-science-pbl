import path from "node:path";
import { pathToFileURL } from "node:url";

const SKILL_DIR = "C:/Users/A/.codex/plugins/cache/openai-primary-runtime/presentations/26.927.11222/skills/presentations";
const WORKSPACE_DIR = "C:/Users/A/ppt_high_dimension_reflection";
const CANDIDATE = path.join(WORKSPACE_DIR, ".codex-build", "student-package-class3-learning-reflection-candidate.pptx");
const FINAL_PPTX = path.join(WORKSPACE_DIR, "out", "student-package-class3-learning-reflection-v8.pptx");
const RUNTIME_PYTHON = "C:/Users/A/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe";

const { finalizePresentation } = await import(
  pathToFileURL(path.join(SKILL_DIR, "container_tools/artifact_tool_utils.mjs")).href,
);

const result = await finalizePresentation({
  workspaceDir: WORKSPACE_DIR,
  candidatePath: CANDIDATE,
  finalPath: FINAL_PPTX,
  pythonExecutable: RUNTIME_PYTHON,
  integrityValidatorPath: path.join(SKILL_DIR, "container_tools/inspect_presentation_package_integrity.py"),
  layoutValidatorPath: path.join(SKILL_DIR, "container_tools/inspect_presentation_layout_geometry.py"),
  layoutArgs: [
    "--expected-slide-size-emu", "12192000,6858000",
    "--expected-slide-count", "11",
    "--cover-role", "cover",
    "--validate-bullet-geometry",
    "--validate-heading-fit",
    "--validate-heading-punctuation",
    "--require-native-table-slide", "4",
  ],
  explicitTotalSlideCount: 11,
  requiredNativeTableOwnerSlides: [4],
  requiredNativeChartOwnerSlides: [5, 6, 7, 8],
  materializeLiteralChartWorkbooks: true,
  fontPolicy: {
    basis: "design",
    families: ["Microsoft YaHei"],
  },
  verifyArtifactToolImport: true,
  receiptPath: path.join(WORKSPACE_DIR, ".codex-build", "student-package-class3-learning-reflection-v8.validation.json"),
});

console.log(JSON.stringify(result, null, 2));
