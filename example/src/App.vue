<template>
  <div class="app">
    <h1>Vue Word Export 示例</h1>
    <p class="subtitle">基于 Vue 3 + TypeScript + docx 的 Word 导出组件</p>

    <!-- ========== 示例1: 企业信用报告 ========== -->
    <section class="demo-section">
      <h2>1. 企业信用报告</h2>
      <p>标题 + 基本信息区(2列) + 股东信息列表 + 对外投资列表</p>
      <WordExport
        :data="creditData"
        :sections="creditSections"
        title="企业信用报告"
        filename="企业信用报告.docx"
        @before-export="onBefore"
        @after-export="onAfter"
        @error="onError"
      >
        <template #trigger="{ exportWord, loading }">
          <button :disabled="loading" class="btn primary" @click="exportWord">
            {{ loading ? '生成中...' : '导出信用报告' }}
          </button>
        </template>
      </WordExport>
    </section>

    <!-- ========== 示例2: 员工信息登记表 ========== -->
    <section class="demo-section">
      <h2>2. 员工信息登记表（单列基本信息）</h2>
      <p>基本信息区使用 columns=1 单列布局，适合字段较多的表单</p>
      <WordExport
        :data="employeeData"
        :sections="employeeSections"
        title="员工信息登记表"
        filename="员工信息登记表.docx"
      >
        <template #trigger="{ exportWord, loading }">
          <button :disabled="loading" class="btn" @click="exportWord">
            {{ loading ? '生成中...' : '导出登记表' }}
          </button>
        </template>
      </WordExport>
    </section>

    <!-- ========== 示例3: 项目进度报告（Hook方式） ========== -->
    <section class="demo-section">
      <h2>3. 项目进度报告（Hook 方式）</h2>
      <p>使用 <code>useWordExporter()</code> 编程式导出，包含自定义格式化</p>
      <button class="btn success" :disabled="hookLoading" @click="handleHookExport">
        {{ hookLoading ? '生成中...' : '导出进度报告' }}
      </button>
    </section>

    <!-- ========== 示例4: 横向页面 + 自定义样式 ========== -->
    <section class="demo-section">
      <h2>4. 横向页面 + 自定义样式</h2>
      <p>pageOrientation=landscape + 自定义表头背景色</p>
      <button class="btn" :disabled="landscapeLoading" @click="handleLandscapeExport">
        {{ landscapeLoading ? '生成中...' : '导出横向报告' }}
      </button>
    </section>

    <!-- ========== 示例5: 多 Section 组合 ========== -->
    <section class="demo-section">
      <h2>5. 复杂报告（多 Section 组合）</h2>
      <p>包含多个基本信息区 + 多个列表区，展示完整文档结构</p>
      <WordExport
        :data="complexData"
        :sections="complexSections"
        title="综合数据分析报告"
        filename="综合报告.docx"
      >
        <template #trigger="{ exportWord, loading }">
          <button :disabled="loading" class="btn danger" @click="exportWord">
            {{ loading ? '生成中...' : '导出综合报告' }}
          </button>
        </template>
      </WordExport>
    </section>
  </div>
</template>

<script setup lang="ts">
import { WordExport, useWordExporter } from 'vue-word-export'
import type { SectionConfig } from 'vue-word-export'

// ── 日志 ─────────────────────────────────────────────────
function onBefore() { console.log('[export] 开始导出') }
function onAfter()  { console.log('[export] 导出完成') }
function onError(e: Error) { console.error('[export] 导出失败', e) }

// ══════════════════════════════════════════════════════════
//  示例1: 企业信用报告
// ══════════════════════════════════════════════════════════

const creditSections: SectionConfig[] = [
  {
    type: 'basic',
    title: '基本信息',
    columns: 2,
    fields: [
      { label: '企业名称', field: 'companyName' },
      { label: '统一社会信用代码', field: 'creditCode' },
      { label: '法定代表人', field: 'legalPerson' },
      { label: '注册资本（万元）', field: 'registeredCapital',
        format: (v) => `${v} 万元` },
      { label: '成立日期', field: 'establishDate' },
      { label: '企业状态', field: 'status' },
      { label: '注册地址', field: 'address' },
      { label: '经营范围', field: 'businessScope' },
    ],
  },
  {
    type: 'list',
    title: '股东信息',
    dataField: 'shareholders',
    columns: [
      { label: '股东名称', field: 'name' },
      { label: '出资比例', field: 'ratio',
        format: (v) => `${(Number(v) * 100).toFixed(1)}%` },
      { label: '认缴出资额（万元）', field: 'amount', align: 'right' },
      { label: '出资日期', field: 'date' },
    ],
  },
  {
    type: 'list',
    title: '对外投资',
    dataField: 'investments',
    columns: [
      { label: '被投资企业', field: 'targetName' },
      { label: '投资比例', field: 'ratio',
        format: (v) => `${(Number(v) * 100).toFixed(1)}%` },
      { label: '投资金额（万元）', field: 'amount', align: 'right' },
    ],
  },
]

const creditData = {
  companyName: '某某科技有限公司',
  creditCode: '91110000MA12345678',
  legalPerson: '张三',
  registeredCapital: 1000,
  establishDate: '2015-03-15',
  status: '存续',
  address: '北京市海淀区中关村大街1号',
  businessScope: '计算机软硬件的技术开发、技术咨询、技术服务',
  shareholders: [
    { name: '张三', ratio: 0.6, amount: 600, date: '2015-03-15' },
    { name: '李四', ratio: 0.3, amount: 300, date: '2015-03-15' },
    { name: '王五', ratio: 0.1, amount: 100, date: '2018-06-01' },
  ],
  investments: [
    { targetName: '北京某某科技发展有限公司', ratio: 0.51, amount: 255 },
    { targetName: '上海某某信息技术有限公司', ratio: 0.3, amount: 150 },
  ],
}

// ══════════════════════════════════════════════════════════
//  示例2: 员工信息登记表（单列）
// ══════════════════════════════════════════════════════════

const employeeSections: SectionConfig[] = [
  {
    type: 'basic',
    title: '个人信息',
    columns: 1,
    fields: [
      { label: '姓名', field: 'name' },
      { label: '性别', field: 'gender' },
      { label: '出生日期', field: 'birthday' },
      { label: '手机号码', field: 'phone' },
      { label: '邮箱', field: 'email' },
      { label: '入职日期', field: 'hireDate' },
      { label: '部门', field: 'department' },
      { label: '职位', field: 'position' },
      { label: '学历', field: 'education' },
      { label: '毕业院校', field: 'school' },
    ],
  },
  {
    type: 'list',
    title: '工作经历',
    dataField: 'workHistory',
    columns: [
      { label: '公司名称', field: 'company' },
      { label: '职位', field: 'position' },
      { label: '起止时间', field: 'period' },
    ],
  },
]

const employeeData = {
  name: '赵六',
  gender: '女',
  birthday: '1995-08-20',
  phone: '138****5678',
  email: 'zhaoliu@company.com',
  hireDate: '2020-03-01',
  department: '产品研发部',
  position: '高级前端工程师',
  education: '本科',
  school: '华中科技大学',
  workHistory: [
    { company: 'ABC科技有限公司', position: '前端工程师', period: '2018.07 - 2020.02' },
    { company: '某某科技有限公司', position: '高级前端工程师', period: '2020.03 - 至今' },
  ],
}

// ══════════════════════════════════════════════════════════
//  示例3: Hook 方式导出
// ══════════════════════════════════════════════════════════

const { exportWord: hookExport, loading: hookLoading } = useWordExporter()

async function handleHookExport() {
  await hookExport({
    filename: '项目进度报告.docx',
    title: '项目进度报告',
    sections: [
      {
        type: 'basic',
        title: '项目概览',
        fields: [
          { label: '项目名称', field: 'name' },
          { label: '项目经理', field: 'manager' },
          { label: '开始日期', field: 'startDate' },
          { label: '计划完成', field: 'endDate',
            format: (v) => v ? String(v) : '待定' },
          { label: '总体进度', field: 'progress',
            format: (v) => `${v}%` },
        ],
      },
      {
        type: 'list',
        title: '任务明细',
        dataField: 'tasks',
        columns: [
          { label: '任务名称', field: 'name' },
          { label: '负责人', field: 'assignee' },
          { label: '状态', field: 'status',
            format: (v) => {
              const map: Record<string, string> = {
                done: '已完成', doing: '进行中', todo: '待开始',
              }
              return map[String(v)] || String(v)
            },
          },
          { label: '完成度', field: 'progress',
            format: (v) => `${v}%`, align: 'right' },
        ],
      },
    ],
    data: {
      name: '企业管理系统重构',
      manager: '陈经理',
      startDate: '2026-01-15',
      endDate: null,
      progress: 65,
      tasks: [
        { name: '需求分析', assignee: '产品组', status: 'done', progress: 100 },
        { name: 'UI 设计', assignee: '设计组', status: 'done', progress: 100 },
        { name: '后端开发', assignee: '后端组', status: 'doing', progress: 70 },
        { name: '前端开发', assignee: '前端组', status: 'doing', progress: 50 },
        { name: '测试部署', assignee: '测试组', status: 'todo', progress: 0 },
      ],
    },
  })
}

// ══════════════════════════════════════════════════════════
//  示例4: 横向页面 + 自定义样式
// ══════════════════════════════════════════════════════════

const { exportWord: landscapeExport, loading: landscapeLoading } = useWordExporter()

async function handleLandscapeExport() {
  await landscapeExport({
    filename: '销售业绩报表（横向）.docx',
    title: {
      text: '2026年第一季度销售业绩报表',
      style: { fontName: '微软雅黑', fontSize: 24, bold: true },
    },
    pageOrientation: 'landscape',
    sections: [
      {
        type: 'basic',
        title: '汇总数据',
        fields: [
          { label: '统计周期', field: 'period' },
          { label: '总销售额（万元）', field: 'totalSales',
            format: (v: unknown) => `${Number(v).toLocaleString()} 万元` },
          { label: '同比增长', field: 'growth',
            format: (v: unknown) => `${v}%` },
          { label: '完成率', field: 'completionRate',
            format: (v: unknown) => `${v}%` },
        ],
      },
      {
        type: 'list',
        title: '各团队销售明细',
        dataField: 'teams',
        columns: [
          { label: '团队', field: 'name' },
          { label: '负责人', field: 'leader' },
          { label: '1月', field: 'jan', align: 'right',
            format: (v: unknown) => `${Number(v).toLocaleString()}` },
          { label: '2月', field: 'feb', align: 'right',
            format: (v: unknown) => `${Number(v).toLocaleString()}` },
          { label: '3月', field: 'mar', align: 'right',
            format: (v: unknown) => `${Number(v).toLocaleString()}` },
          { label: '合计', field: 'total', align: 'right',
            format: (v: unknown) => `${Number(v).toLocaleString()}` },
        ],
      },
    ],
    data: {
      period: '2026年1月 - 3月',
      totalSales: 2850,
      growth: 18.5,
      completionRate: 95,
      teams: [
        { name: '华东团队', leader: '张经理', jan: 320, feb: 350, mar: 380, total: 1050 },
        { name: '华南团队', leader: '李经理', jan: 280, feb: 310, mar: 360, total: 950 },
        { name: '华北团队', leader: '王经理', jan: 200, feb: 250, mar: 280, total: 730 },
        { name: '西南团队', leader: '赵经理', jan: 150, feb: 170, mar: 200, total: 520 },
      ],
    },
  })
}

// ══════════════════════════════════════════════════════════
//  示例5: 复杂多 Section 组合
// ══════════════════════════════════════════════════════════

const complexSections: SectionConfig[] = [
  {
    type: 'basic',
    title: '报告概述',
    fields: [
      { label: '报告编号', field: 'reportId' },
      { label: '生成日期', field: 'date' },
      { label: '报告类型', field: 'type' },
      { label: '数据截止', field: 'cutoffDate' },
    ],
  },
  {
    type: 'basic',
    title: '核心指标',
    columns: 2,
    fields: [
      { label: '总用户数', field: 'metrics.totalUsers',
        format: (v: unknown) => `${Number(v).toLocaleString()}` },
      { label: '活跃用户', field: 'metrics.activeUsers',
        format: (v: unknown) => `${Number(v).toLocaleString()}` },
      { label: '总订单数', field: 'metrics.totalOrders',
        format: (v: unknown) => `${Number(v).toLocaleString()}` },
      { label: '成交金额（万元）', field: 'metrics.revenue',
        format: (v: unknown) => `${Number(v).toLocaleString()}` },
      { label: '转化率', field: 'metrics.conversionRate',
        format: (v: unknown) => `${v}%` },
      { label: '客单价（元）', field: 'metrics.avgOrderValue',
        format: (v: unknown) => `¥${v}` },
    ],
  },
  {
    type: 'list',
    title: '月度趋势',
    dataField: 'monthlyTrend',
    columns: [
      { label: '月份', field: 'month' },
      { label: '新用户', field: 'newUsers', align: 'right' },
      { label: '订单数', field: 'orders', align: 'right' },
      { label: '营收（万元）', field: 'revenue', align: 'right',
        format: (v) => `${v}` },
    ],
  },
  {
    type: 'list',
    title: '产品分类统计',
    dataField: 'categories',
    columns: [
      { label: '产品类别', field: 'name' },
      { label: '销量', field: 'sales', align: 'right' },
      { label: '销售额（万元）', field: 'revenue', align: 'right',
        format: (v) => `${(v as number).toFixed(1)}` },
      { label: '占比', field: 'share',
        format: (v) => `${(v as number).toFixed(1)}%` },
    ],
  },
]

const complexData = {
  reportId: 'REP-2026-0001',
  date: '2026-06-01',
  type: '月度运营报告',
  cutoffDate: '2026-05-31',
  metrics: {
    totalUsers: 158000,
    activeUsers: 82300,
    totalOrders: 45600,
    revenue: 3280,
    conversionRate: 3.2,
    avgOrderSize: 719,
    avgOrderValue: 719,
  },
  monthlyTrend: [
    { month: '2026-01', newUsers: 8500, orders: 3800, revenue: 285 },
    { month: '2026-02', newUsers: 9200, orders: 4200, revenue: 310 },
    { month: '2026-03', newUsers: 10500, orders: 4800, revenue: 356 },
    { month: '2026-04', newUsers: 11800, orders: 5100, revenue: 382 },
    { month: '2026-05', newUsers: 12600, orders: 5500, revenue: 412 },
  ],
  categories: [
    { name: '电子产品', sales: 12800, revenue: 1280, share: 39.0 },
    { name: '服装鞋帽', sales: 15600, revenue: 780, share: 23.8 },
    { name: '食品饮料', sales: 8900, revenue: 445, share: 13.6 },
    { name: '家居用品', sales: 5300, revenue: 375, share: 11.4 },
    { name: '其他', sales: 3000, revenue: 400, share: 12.2 },
  ],
}
</script>

<style>
body {
  margin: 0;
  background: #f5f6fa;
}
.app {
  max-width: 960px;
  margin: 0 auto;
  padding: 2rem;
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', sans-serif;
  color: #333;
}
h1 {
  margin: 0 0 0.25rem;
  font-size: 1.75rem;
  color: #1a1a2e;
}
.subtitle {
  margin: 0 0 2rem;
  color: #888;
  font-size: 0.9rem;
}
.demo-section {
  background: #fff;
  border: 1px solid #e8ecf1;
  border-radius: 10px;
  padding: 1.5rem;
  margin-bottom: 1.5rem;
  box-shadow: 0 1px 3px rgba(0,0,0,0.04);
}
.demo-section h2 {
  margin: 0 0 0.25rem;
  font-size: 1.05rem;
  color: #2d3436;
}
.demo-section p {
  margin: 0 0 1rem;
  font-size: 0.85rem;
  color: #999;
}
.btn {
  padding: 0.55rem 1.5rem;
  border: 1px solid #dfe6e9;
  border-radius: 6px;
  background: #fff;
  cursor: pointer;
  font-size: 0.9rem;
  transition: all 0.15s;
}
.btn:hover { border-color: #b2bec3; }
.btn:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}
.btn.primary {
  background: #0984e3;
  color: #fff;
  border-color: #0984e3;
}
.btn.primary:hover { background: #0773c5; }
.btn.success {
  background: #00b894;
  color: #fff;
  border-color: #00b894;
}
.btn.success:hover { background: #00a381; }
.btn.danger {
  background: #d63031;
  color: #fff;
  border-color: #d63031;
}
.btn.danger:hover { background: #c0392b; }
code {
  background: #f1f2f6;
  padding: 0.1em 0.35em;
  border-radius: 3px;
  font-size: 0.88em;
  color: #636e72;
}
</style>
