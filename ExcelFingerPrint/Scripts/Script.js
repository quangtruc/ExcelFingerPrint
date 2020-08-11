var Home = {
    Init() {
        return this.Register();
    },
    Register() {
        $('.select-2').select2();
        $('.select-2-multiple').select2({
            placeholder: 'Select Entry Door'
        });

        let self = this;

        $('#upload-file').click(function () {
            let fileUpload = $("#file-upload").get(0);
            let file = fileUpload.files;

            // Create FormData object  
            let fileData = new FormData();
            fileData.append("file", file[0]);

            self.Methods.ImportExcel(fileData);
        });

        $('#search').click(function () {
            let guestID = $('#guest-id').val();
            let listEntryDoor = $('#list-entry-door').val();
            if (listEntryDoor.length == 0 && guestID == '') {
                alert('Vui lòng chọn cửa vào hoặc guestID để tìm');
                return;
            }
            self.Methods.LoadData(guestID, listEntryDoor);
        });

        $('#export-excel').click(function () {
            self.Methods.ExpotExcel();
        });

    },
    Methods: {
        ShowLoading() {
            $("#main-loading").fadeIn();
        },
        HideLoading() {
            $("#main-loading").fadeOut(500);
        },
        LoadData(guestID, listEntryDoor) {
            this.ShowLoading();
            $.post('/Home/Search', { guestID: guestID, listEntryDoor: listEntryDoor }).then(res => {
                this.HideLoading();
                $('#data-finger-print').html(res);
                $('#guest-id').val() = "";
            });

        },
        ImportExcel(fileData) {
            let seft = this;
            this.ShowLoading();

            $.ajax({
                url: '/Home/ImportExcel',
                type: 'POST',
                data: fileData,
                processData: false,  // tell jQuery not to process the data
                contentType: false,  // tell jQuery not to set contentType
                success: function (res) {
                    if (res.status == true) {
                        alert('Success \n ' + res.message);
                    }
                    else {
                        alert('Failed \n ' + res.message);
                    }
                    window.location.href = '/Home/Index';
                    seft.HideLoading();
                }
            });
            //$.post('/Home/ImportExcel', a).then(res => {
            //    if (res.status == true) {
            //        alert('Success \n ' + res.message);
            //    }
            //    else {
            //        alert('Failed \n ' + res.message);
            //    }
            //    this.LoadData('');
            //    this.HideLoading();
            //});
        },
        ExpotExcel() {
            window.location.href = '/Home/ExportExcel';
        },
    },
    Helpers() {

    }
}
Home.Init();