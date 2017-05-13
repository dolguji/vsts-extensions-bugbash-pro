import * as WitClient from "TFS/WorkItemTracking/RestClient";
import { AttachmentReference } from "TFS/WorkItemTracking/Contracts";

(function ($) {
    'use strict';

    $.extend(true, ($ as any).trumbowyg, {
        plugins: {
            pasteImage: {
                init: function (trumbowyg) {
                    trumbowyg.pasteHandlers.push(function (pasteEvent) {
                        try {
                            var items = (pasteEvent.originalEvent || pasteEvent).clipboardData.items,
                                reader;

                            for (var i = items.length -1; i >= 0; i += 1) {
                                if (items[i].type.indexOf("image") === 0) {
                                    reader = new FileReader();
                                    reader.onloadend = async (event) => {
                                        const fileName = `pastedimage_${Date.now().toString()}.png`;
                                        const attachment = await WitClient.getClient().createAttachment(event.target.result, fileName);
                                        trumbowyg.execCmd('insertImage', attachment.url, undefined, true);
                                    };
                                    reader.readAsArrayBuffer(items[i].getAsFile());
                                    break;
                                }
                            }
                        } catch (c) {
                        }
                    });
                }
            }
        }
    });
})(jQuery);
