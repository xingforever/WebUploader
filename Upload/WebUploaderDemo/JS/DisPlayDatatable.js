$(function () {
    $("#ShowExcel").click(function () {
        var pars = {
            Adress: $("#ExcelAdress").val(),//excel地址

        };

        loadData(pars);

    })

    function loadData(pars) {
        $('#dg').datagrid({
            url: '/DisplayExcel/GetCityInfoList',
            //  width: 'auto',
            //  height: 'auto',
            title: '中国城市',
            fitColumns: true, //列自适应
            nowrap: false,
            idField: 'Column1',//主键列的列明
            loadMsg: '正在加载中国城市的信息...',
            striped: true, //是否显示斑马线
            pagination: true,//是否有分页
            singleSelect: false,//是否单行选择
            pageSize: 5,//页大小，一页多少条数据
            pageNumber: 1,//当前页，默认的
            toolbar: "#tb",//工具条        
            fit: true,
            pageList: [5, 10, 15],
            queryParams: pars,//往后台传递参数
            columns: [[//c.UserName, c.UserPass, c.Email, c.RegTime
                { field: 'ck', checkbox: true, align: 'left', width: 50 },
                { field: 'Column1', title: '地区ID', width: 200 },
                { field: 'Column2', title: '地区代码', width: 200 },
                { field: 'Column3', title: '地区名称', width: 100 },
                { field: 'Column4', title: '父级ID', width: 100 },
                { field: 'Column5', title: '地区水平', width: 100 },
                { field: 'Column6', title: '地区等级', width: 100 },
                { field: 'Column7', title: '英文名称', width: 100 },
                { field: 'Column8', title: '英文简称', width: 100 },
            ]]


        });
    }

})

//columns: [[//c.UserName, c.UserPass, c.Email, c.RegTime
//    { field: 'ck', checkbox: true, align: 'left', width: 50 },
//    { field: 'REGION_ID', title: '地区ID', width: 200 },
//    { field: 'REGION_CODE', title: '地区代码', width: 200 },
//    { field: 'REGION_NAME', title: '地区名称', width: 100 },
//    { field: 'PARENT_ID', title: '父级ID', width: 100 },
//    { field: 'REGION_LEVEL', title: '地区水平', width: 100 },
//    { field: 'REGION_ORDER', title: '地区等级', width: 100 },
//    { field: 'REGION_NAME_EN', title: '英文名称', width: 100 },
//    { field: 'REGION_SHORTNAME_EN', title: '英文简称', width: 100 },
//]]