(function ($) {
    'use strict';

    $.extend(true, ($ as any).trumbowyg, {
        plugins: {
            pasteImage: {
                init: function (trumbowyg) {
                    trumbowyg.pasteHandlers.push(function (pasteEvent) {
                        try {
                            const items = (pasteEvent.originalEvent || pasteEvent).clipboardData.items;
                            
                            for (var i = items.length -1; i >= 0; i += 1) {
                                if (items[i].type.indexOf("image") === 0) {
                                    const reader = new FileReader();
                                    reader.onloadend = (event) => {
                                        const data = (event.target as any).result;

                                        $(window).trigger("imagepasted", { data: data, callback: (imageUrl: string) => {
                                            if (imageUrl) {
                                                trumbowyg.execCmd("insertImage", imageUrl, undefined, true);
                                                $(['img[src="' + imageUrl + '"]:not([alt])'].join(''), trumbowyg.$box).css("width", "auto").css("height", "400px");
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