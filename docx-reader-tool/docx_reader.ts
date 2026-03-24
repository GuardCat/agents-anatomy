// place it to ~/.config/opencode/tools/docx_reader.ts

import { tool } from "@opencode-ai/plugin"
import { exec } from "node:child_process"
import { promisify } from "node:util"

const execPromise = promisify(exec)

export default tool({
  // Это описание пойдет прямо в системный промпт агента. 
  // Пиши максимально четко, чтобы LLM понимала, когда это использовать.
  description: "Reads and extracts text, tables, inline links, and images from a Word (.docx) document. Use this INSTEAD of the standard 'read' tool for any .docx files.",
  
  // Строгая типизация аргументов с помощью Zod
  args: {
    file_path: tool.schema.string().describe("Absolute path to the .docx file"),
  },
  
  async execute(args) {
    try {
      // Вызываем наш глобальный bash-алиас, который дергает uv и python-docx
      const { stdout, stderr } = await execPromise(`/usr/local/bin/docx_reader.py "${args.file_path}"`);
      
      if (stderr) {
        console.warn("Tool warning:", stderr);
      }
      
      // Возвращаем готовый Markdown прямо в контекст агента
      return stdout;
    } catch (error) {
      return `Ошибка чтения DOCX: ${error.message}`;
    }
  }
})
