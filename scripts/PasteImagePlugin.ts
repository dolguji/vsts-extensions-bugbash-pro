(function ($) {
    'use strict';

    $.extend(true, ($ as any).trumbowyg, {
        plugins: {
            pasteImage: {
                init: function (trumbowyg) {
                    trumbowyg.pasteHandlers.push(function (pasteEvent) {
                        try {
                            let items = (pasteEvent.originalEvent || pasteEvent).clipboardData.items, reader;
                            
                            for (var i = items.length -1; i >= 0; i += 1) {
                                if (items[i].type.indexOf("image") === 0) {
                                    reader = new FileReader();
                                    reader.onloadend = (event) => {
                                        const data = event.target.result;

                                        $(window).trigger("imagepasted", { data: data, callback: (imageUrl: string) => {
                                            if (imageUrl) {
                                                trumbowyg.execCmd('insertImage', imageUrl, undefined, true);
                                            }
                                        }});
                                    };
                                    reader.readAsDataURL(items[i].getAsFile());
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