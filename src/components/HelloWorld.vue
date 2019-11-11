<template>
  <div class="hello">
  </div>
</template>

<script>
import { handleExport } from '@/utils/public'
export default {
  name: 'HelloWorld',
  data () {
    return {
      anchorTable: [
        {
          classBook: "",
          className: "",
          endTime: "",
          errorDescribe: "教师账号已注册",
          organName: "",
          startTime: "",
          studentName: "",
          studentNo: "",
          teacherName: "fghfdh",
          teacherOrgan: "苏州新东方",
          teacherPhone: "13245455345"
        },
        {
          classBook: "",
          className: "",
          endTime: "",
          errorDescribe: "教师账号已注册",
          organName: "",
          startTime: "",
          studentName: "",
          studentNo: "",
          teacherName: "fghfdh",
          teacherOrgan: "苏州新东方",
          teacherPhone: "13245455345"
        },
        {
          classBook: "",
          className: "",
          endTime: "",
          errorDescribe: "教师账号已注册",
          organName: "",
          startTime: "",
          studentName: "",
          studentNo: "",
          teacherName: "fghfdh",
          teacherOrgan: "苏州新东方",
          teacherPhone: "13245455345"
        },
        {
          classBook: "",
          className: "",
          endTime: "",
          errorDescribe: "教师账号已注册",
          organName: "",
          startTime: "",
          studentName: "",
          studentNo: "",
          teacherName: "xcvbxv",
          teacherOrgan: "苏州新东方",
          teacherPhone: "13453455355"
        }
      ]
    }
  },
  mounted () {
    this.exportExcel()
  },
  methods: {
    exportExcel() {
      //titleList,mergeList 需要合并单元格和加表头的时候，可以加上这两个参数，格式 titleList = ['名字','','','',''],titleList数组的长度和表格对应字段一样长，只是第一个有值，其余都是空,mergeList = ['A1:D1']
      let tHeader = ['教师的错误运营', '教师的姓名', '教师的所属机构', '教师的手机号', '备注'], //表头
          fileName = '示例DEMO的表格名字',  //标题
          filterVal = ['errorDescribe','teacherName','teacherOrgan','teacherPhone','organName'],//表头对应字段
          dataList = this.anchorTable
      handleExport(dataList,tHeader,fileName,filterVal)
    }
  }
}
</script>

<!-- Add "scoped" attribute to limit CSS to this component only -->
<style scoped>
h1, h2 {
  font-weight: normal;
}
ul {
  list-style-type: none;
  padding: 0;
}
li {
  display: inline-block;
  margin: 0 10px;
}
a {
  color: #42b983;
}
</style>
