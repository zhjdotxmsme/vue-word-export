<template>
  <slot
    name="trigger"
    :exportWord="handleExport"
    :loading="loading"
  >
    <button
      :disabled="loading"
      @click="handleExport"
    >
      {{ loading ? '导出中...' : '导出 Word' }}
    </button>
  </slot>
</template>

<script setup lang="ts">
import { watch } from 'vue'
import type { SectionConfig, TextStyle, ExportOptions } from '../types'
import { useWordExporter } from '../composables/useWordExporter'

export interface WordExportProps {
  /** 文档标题（可传入字符串或配置对象） */
  title?: string | {
    text: string | (() => string)
    style?: Partial<TextStyle>
  }
  /** 文档 Section 配置 */
  sections: SectionConfig[]
  /** 后端 JSON 数据 */
  data: Record<string, unknown>
  /** 导出文件名（默认 'export.docx'） */
  filename?: string
  /** 页面方向 */
  pageOrientation?: 'portrait' | 'landscape'
  /** 页边距（twips） */
  margins?: {
    top?: number
    bottom?: number
    left?: number
    right?: number
  }
}

const props = withDefaults(defineProps<WordExportProps>(), {
  filename: 'export.docx',
  pageOrientation: 'portrait',
})

const emit = defineEmits<{
  (e: 'before-export'): void
  (e: 'after-export'): void
  (e: 'error', err: Error): void
}>()

const { exportWord, loading, error } = useWordExporter()

watch(error, (err) => {
  if (err) {
    emit('error', err)
  }
})

async function handleExport() {
  emit('before-export')
  try {
    const normalizedTitle = typeof props.title === 'string'
      ? { text: props.title }
      : props.title

    await exportWord({
      title: normalizedTitle,
      sections: props.sections,
      data: props.data,
      filename: props.filename,
      pageOrientation: props.pageOrientation,
      margins: props.margins,
    })
    emit('after-export')
  } catch {
    // error already handled by useWordExporter
  }
}
</script>
