import * as ImageCompressor from '@xkeshi/image-compressor';

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
                                                $(['img[src="' + imageUrl + '"]:not([alt])'].join(''), trumbowyg.$box).css("width", "auto").css("max-height", "400px");
                                            }
                                        }});
                                    };

                                    const file = items[i].getAsFile();
                                    new ImageCompressor(file, {
                                        quality: .6,
                                        convertSize: 2000000,
                                        success(newFile) {
                                            reader.readAsDataURL(newFile);
                                        },
                                        error(e) {
                                            console.log(e.message);
                                        },
                                    });
                                    
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