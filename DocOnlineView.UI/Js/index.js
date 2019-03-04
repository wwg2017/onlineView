
$(function () {
    SexyLightbox.initialize({ color: 'white' });
});

var viewDoc = function (fileName) {
    showLoading("body", "正在生成预览");

    //生成html文件  
    $.ajax({
        url: "api/Home/CourseViewOnLine?fileName=" + fileName,
        type: "GET",
        dataType: "json",
        success: function (data) {
            closeLoading();
 
            //alert(JSON.stringify(data));
            //alert(data[0].TempDocHtml);
            var diag = new Dialog();
            diag.Width = 900;
            diag.Height = 400;
            diag.Title = "内容页为外部连接的窗口";
            diag.URL = "http://localhost:8640/" + data[0].TempDocHtml + "?ver=" + Math.random() * 10;
            diag.show();

            //$("#hidePopupDialog").attr('href', '' + data[0].TempDocHtml + '?TB_iframe=true&height=450&width=920');
            //$("#hidePopupDialog").click();
        },
        error: function () {
            closeLoading();
            alert('生成失败');
        }
    });
}

// 加载遮罩
var showLoading = function (elementTag, message) {
    var msg = message ? message : "加载数据，请稍候...";
    $("<div class=\"datagrid-mask\"></div>").css({
        display: "block", width: "100%",
        height: $(elementTag).height()
    }).appendTo(elementTag);
    $("<div class=\"datagrid-mask-msg\"></div>")
        .html(msg)
        .appendTo(elementTag).css({ display: "block", left: "30%", top: ($(elementTag).height() - 45) / 2 });
};

//关闭遮罩
var closeLoading = function () {
    $('.datagrid-mask').remove();
    $('.datagrid-mask-msg').remove();
};