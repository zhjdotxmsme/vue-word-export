import { ref, shallowRef } from 'vue'
import type { ExportOptions } from '../types/export'
import { exportWord as engineExport, buildDocument } from '../engine/builder'

/**
 * Vue3 Hook: Word 文档导出功能
 *
 * @example
 * ```ts
 * const { exportWord, loading, error } = useWordExporter()
 *
 * await exportWord({
 *   filename: '报告.docx',
 *   sections: [...],
 *   data: response.data,
 * })
 * ```
 */
export function useWordExporter() {
  const loading = ref(false)
  const error = shallowRef<Error | null>(null)

  /**
   * 导出 Word 文件（构建 + 下载）
   */
  async function exportWord(options: ExportOptions): Promise<void> {
    loading.value = true
    error.value = null

    try {
      await engineExport(options)
    } catch (err) {
      const e = err instanceof Error ? err : new Error(String(err))
      error.value = e
      throw e
    } finally {
      loading.value = false
    }
  }

  /**
   * 仅构建文档 Blob（不触发下载）
   */
  async function build(options: ExportOptions): Promise<Blob> {
    loading.value = true
    error.value = null

    try {
      return await buildDocument(options)
    } catch (err) {
      const e = err instanceof Error ? err : new Error(String(err))
      error.value = e
      throw e
    } finally {
      loading.value = false
    }
  }

  return {
    /** 导出 Word（构建 + 下载） */
    exportWord,
    /** 仅构建文档，返回 Blob */
    build,
    /** 是否正在导出 */
    loading,
    /** 导出错误 */
    error,
  }
}
