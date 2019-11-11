<template>
  <div class="hello">
    <el-upload
        class="upload-demo"
        action="https://interface.xuexi8.net/jiekou2/index.php?c=FunLearning&a=uploadFileToAliOSS"
        :on-change="handleOnchange"
        :file-list="fileList"
        :multiple="false"
        :before-upload="beforeAvatarUpload"
        :on-success="handleSuccess"
        :data="myData">
      <el-button size="small" type="primary">点击上传</el-button>
    </el-upload>
  </div>
</template>

<script>
export default {
  name: 'HelloWorld',
  data () {
    return {
      msg: 'Welcome to Your Vue.js App',
      fileList: [],
      myData: {
        fileName: 'file',
        fileType: 1,
        fileSize: 0
      },
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
    handleOnchange(file, fileList) {
      this.fileList = fileList.slice();
    },
    handleSuccess() {
      this.$message({
        message: "上传成功！",
        type: "success"
      });
    },
    async beforeAvatarUpload(file) {
      console.log(file)
      let fileType = file.name.split('.')[1]
      console.log(fileType)
      this.myData.fileSize = file.size
      if (fileType == 'png' || fileType == 'jpg' || fileType == 'jpeg') {
        this.myData.fileType = 1
      } else if (fileType == 'mp3' || fileType == 'flac') {
        this.myData.fileType = 2
      } else if (fileType == 'mp4' || fileType == 'mkv') {
        this.myData.fileType = 3
      } else if (fileType == 'docx' || fileType == 'xlsx' || fileType == 'pdf') {
        this.myData.fileType = 4
      }
    },
    exportExcel() {
      import('@/vendor/excelOut').then(excel => {
        const tHeader = ['教师的错误运营', '教师的姓名', '教师的所属机构', '教师的手机号', '备注'] //表头
        const title = ['锚地船舶', '', '', '', '']  //标题
      //表头对应字段
        const filterVal = ['errorDescribe','teacherName','teacherOrgan','teacherPhone','organName']
        const list = this.anchorTable
        const data = this.formatJson(filterVal, list)
        data.map(item => {
          // console.log(item)
          item.map((i, index) => {
            if (!i) {
              item[index] = ''
            }
          })
        })
        // 空数组和非空数组，出来的表格是不一样的，都是可以自己定制的
        // const merges = ['A1:E1'] //合并单元格的参数，excel表格，分横向是字母A-Z，纵向是数字1-很多，所以A1就代表第一个格子
        const merges = []
        excel.exportJsonToExcel({
          title: title,
          header: tHeader,
          data,
          merges,
          filename: '锚地船舶',
          autoWidth: true,
          bookType: 'xlsx',
          myRowFont: '2'
        })
      })
    },
    formatJson(filterVal, jsonData) {
      return jsonData.map(v => filterVal.map(j => v[j]))
    },
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
