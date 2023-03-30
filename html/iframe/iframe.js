/**
 *
 * (c) copyright ascensio system sia 2020
 *
 * licensed under the apache license, version 2.0 (the "license");
 * you may not use this file except in compliance with the license.
 * you may obtain a copy of the license at
 *
 *     http://www.apache.org/licenses/license-2.0
 *
 * unless required by applicable law or agreed to in writing, software
 * distributed under the license is distributed on an "as is" basis,
 * without warranties or conditions of any kind, either express or implied.
 * see the license for the specific language governing permissions and
 * limitations under the license.
 *
 */
window.onload = function () {

    var datamessage = {
        frameeditorid: "iframeeditor",
        guid: "asc.{a8705dee-7544-4c33-b3d5-168406d92f72}",
        type: "onexternalpluginmessage",
        data: {
            type: "close",
            text: "text"
        }
    };

    document.getElementById("button1").onclick = () => {
        datamessage.data.type = "inserttext";
        datamessage.data.text = "text1";
        var _iframe = document.getElementById("iframeeditor");
        if (_iframe)
            _iframe.contentwindow.postMessage(json.stringify(datamessage), "*");
    };
    document.getElementById("button2").onclick = () => {
        datamessage.data.type = "inserttext";
        datamessage.data.text = "text2";
        var _iframe = document.getElementById("iframeeditor");
        if (_iframe)
            _iframe.contentwindow.postMessage(json.stringify(datamessage), "*");
    };
    document.getElementById("buttonclose").onclick = function () {
        datamessage.data.type = "close";
        datamessage.data.text = "";
        var _iframe = document.getElementById("iframeeditor");
        if (_iframe)
            _iframe.contentwindow.postMessage(json.stringify(datamessage), "*");
    };
};
