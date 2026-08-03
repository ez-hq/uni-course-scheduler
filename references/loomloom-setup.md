# LoomLoom CLI Setup Guide

## What is LoomLoom?

LoomLoom is a batch LLM cloud execution platform. It runs structured tasks in
parallel and returns structured results. This skill uses LoomLoom for the
3-step cloud pipeline:

1. **stp_catalog** — Course Catalog Analysis
2. **stp_recommend** — Course Recommendation
3. **stp_schedule** — Weekly Schedule Generation

## Installation

### Option A: One-Click Install

Paste this prompt to any AI assistant:

> 请你在这个项目里安装 LoomLoom：安装地址是
> https://github.com/Cogfoundry-ai/loomloom 或
> https://gitee.com/cogfoundry/loomloom
> 服务地址：https://loomloom.shengsuanyun.com/loom/v1
> Token 是【替换成你自己的胜算云密钥】。
> 安装好之后帮我运行一次 doctor 检查是否正常。

### Option B: Manual Install

1. Apply for an API key at https://console.shengsuanyun.com/user/keys
2. Clone the repository:
   ```bash
   git clone https://gitee.com/cogfoundry/loomloom.git
   ```
3. Run the install script inside the cloned directory (follow the repo README)
4. Configure environment variables in `~/.zshrc` or `~/.bashrc`:
   ```bash
   export LOOMLOOM_SERVER='https://loomloom.shengsuanyun.com/loom/v1'
   export LOOMLOOM_TOKEN='your-shengsuanyun-api-key'
   ```
5. Verify:
   ```bash
   source ~/.zshrc
   loomloom doctor
   ```

## Without LoomLoom

If LoomLoom is not installed, this skill still works in **local-only mode**.
The Agent uses its own reasoning to:
- Analyze the course catalog (via web search)
- Generate recommendations based on the student's goals
- Create a weekly schedule with conflict detection

The Excel generation and validation scripts work identically in both modes.

## Cloud Execution Flow

When LoomLoom is available:

1. Validate the template-spec:
   ```bash
   loomloom template-spec check template-spec.json
   ```

2. Create the template:
   ```bash
   loomloom template-spec create template-spec.json
   ```

3. Download the input workbook, fill it with the student's data.

4. Validate and precheck:
   ```bash
   loomloom template-spec validate-workbook <template-id> <version-id> input.xlsx
   loomloom template-spec precheck-workbook <template-id> <version-id> input.xlsx
   ```

5. Show the fee estimate to the student and get confirmation.

6. Execute:
   ```bash
   loomloom template-spec submit-workbook <template-id> <version-id> input.xlsx \
     --client-request-id <unique-id>
   ```

7. Watch and retrieve results:
   ```bash
   loomloom run watch <run-id>
   loomloom run result-rows <run-id>
   ```

8. Pass the combined JSON to `scripts/generate_excel.py`.
