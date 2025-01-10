
// Example insert text into editors (different implementations)
(function(window, undefined){

    window.Asc.plugin.init = function()
    {
        console.info('1')

        let file = 'index.html'
        let variation = {
            url : location.href.replace(file, 'modal.html'),
            description : window.Asc.plugin.tr('Warning'),
            isVisual : true,
            isModal : true,
            EditorsSupport : ['word', 'cell', 'slide'],
            size : [350, 100],
            buttons : [
                {
                    'text': window.Asc.plugin.tr('Yes'),
                    'primary': true
                },
                {
                    'text': window.Asc.plugin.tr('No'),
                    'primary': false
                }
            ]
        };

        // DOES NOT WORK
        // function openModal(){
        //     console.log('click')
        //     window.Asc.plugin.executeMethod ("ShowWindow", ["iframe_asc.{BE5CBF95-C0AD-4842-B157-AC40FEDD9841}", variation]);
        // }

        // DOES NOT WORK
        // window.onload = openModal;
        // function say(){
        //     console.log('click')
        // }
        //
        // document.onload = function () {
        //     document.getElementById("01").onclick = say
        // };

        // DOES NOT WORK
        // document.getElementById('open').onclick = openModal()
        // document.getElementById('open').addEventListener('click', openModal());

        // window.onload = function (){
        //     document.getElementById('open').addEventListener('click', function (){
        //         window.Asc.plugin.executeMethod ("ShowWindow", ["iframe_asc.{BE5CBF95-C0AD-4842-B157-AC40FEDD9841}", variation]);
        //     })
        // }

        document.getElementById('open').addEventListener('click', function (){
            window.Asc.plugin.executeMethod ("ShowWindow", ["iframe_asc.{BE5CBF95-C0AD-4842-B157-AC40FEDD9841}", variation]);
        })

        document.getElementById('all_plugins').addEventListener('click', function (){
            window.Asc.plugin.executeMethod ("GetInstalledPlugins", [], function (plugins){
                console.log(plugins)
            });
        })
    };

    // function openModal(){
    //     console.info('2')
    //     let file = 'index.html'
    //     let variation = {
    //         url : location.href.replace(file, 'modal.html'),
    //         description : window.Asc.plugin.tr('Warning'),
    //         isVisual : true,
    //         isModal : true,
    //         EditorsSupport : ['word', 'cell', 'slide'],
    //         size : [350, 100],
    //         buttons : [
    //             {
    //                 'text': window.Asc.plugin.tr('Yes'),
    //                 'primary': true
    //             },
    //             {
    //                 'text': window.Asc.plugin.tr('No'),
    //                 'primary': false
    //             }
    //         ]
    //     };
    //     window.Asc.plugin.executeMethod ("ShowWindow", ["iframe_asc.{BE5CBF95-C0AD-4842-B157-AC40FEDD9841}", variation]);
    // }

    window.Asc.plugin.button = function(id, windowId)
    {
        console.info('3')

        console.log(id)
        console.log(windowId)
        if (typeof(windowId) == 'undefined') {
            this.executeCommand("close", "");
        }

        if (windowId) {
            switch (id) {
                case 0:
                    console.log('yes');
                    break;
                case 1:
                    console.log('no');
                    break;
                default:
                    window.Asc.plugin.executeMethod('CloseWindow', [windowId]);
            }
        }
    };

})(window, undefined);
