using ExcelKit.Core.Attributes;
using ExcelKit.Core.Constraint.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BigDataExcel
{
    class Person
    {
        [ExcelKit(Desc = "用户名", Width = 20, IsIgnore = false, Sort = 10, Align = TextAlign.Center, FontColor = DefineColor.LightBlue)]
        public string UserName { get; set; }
        [ExcelKit(Desc = "密码", Width = 20, Sort = 20, FontColor = DefineColor.Rose)]
        public string Pwd { get; set; }
        [ExcelKit(Desc = "住址", Width = 30, Sort = 30, FontColor = DefineColor.Rose, ForegroundColor = DefineColor.LemonChiffon)]
        public string Address { get; set; }
        [ExcelKit(Desc = "年龄", Width = 10, Sort = 40)]
        public int  Age { get; set; }
        [ExcelKit(Desc = "兴趣爱好", Width = 30, Sort = 50, FontColor = DefineColor.Rose, ForegroundColor = DefineColor.LemonChiffon)]
        public string Hobby { get; set; }

    }
}
