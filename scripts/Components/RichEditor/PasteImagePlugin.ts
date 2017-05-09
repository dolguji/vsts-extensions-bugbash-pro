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
                                if (items[i].type.match(/^image\//)) {
                                    reader = new FileReader();
                                    reader.onloadend = async (event) => {
                                        // const attachment = await WitClient.getClient().createAttachment(event.target.result, "pastedimage.png");
                                        // trumbowyg.execCmd('insertImage', attachment.url, undefined, true);
                                        trumbowyg.execCmd('insertImage', event.target.result, undefined, true);
                                    };
                                    reader.readAsDataURL(items[i].getAsFile());
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
