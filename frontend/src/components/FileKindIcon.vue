<template>
  <span class="file-badge" :class="`is-${kind}`">
    <FileSpreadsheet v-if="kind === 'sheet'" :size="size" />
    <FileText v-else-if="kind === 'word' || kind === 'pdf'" :size="size" />
    <FileImage v-else-if="kind === 'image'" :size="size" />
    <File v-else :size="size" />
  </span>
</template>

<script setup>
import { computed } from 'vue'
import { File, FileImage, FileSpreadsheet, FileText } from 'lucide-vue-next'
import { getFileKind } from '@/utils/format'

const props = defineProps({
  filename: { type: String, default: '' },
  size: { type: Number, default: 17 },
})

const kind = computed(() => getFileKind(props.filename))
</script>
