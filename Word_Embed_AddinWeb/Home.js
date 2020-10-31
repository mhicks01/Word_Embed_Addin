
(function () {
    "use strict";

    var messageBanner;

    var one = "https://www.numworks.com/simulator/";
    var two = "https://www.desmos.com/";
    var three = "https://www.wolframalpha.com/";
    var four = "https://www.wikipedia.org/";
    var five = "https://www.cpp.edu/it/client-services/software/software-students.shtml";
    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the notification mechanism and hide it
            var element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();

            // If not using Word 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('WordApi', '1.1')) {
                $('#button-text').text("Display!");
                $('#button-desc').text("Display the selected text");
                $('#webOne-button').click(displaySelectedText);
                return;
            }



            $('#button-textOne').text("Numworks");
            $('#button-descOne').text("Take you to Numworks website.");

            $('#button-textTwo').text("Desmos");
            $('#button-descTwo').text("Take you to Desmos website.");

            $('#button-textThree').text("Wolframalpha");
            $('#button-descThree').text("Take you to Wolframalpha website.");

            $('#button-textFour').text(" Wikipedia ");
            $('#button-descFour').text("Take you to Wikipedia.");

            $('#button-textFive').text("CPP Software");
            $('#button-descFive').text("Take you to the CPP software page.");

            // Add a click event handler for the highlight button.
            $('#webOne-button').click(webOne);
            $('#webTwo-button').click(webTwo);
            $('#webThree-button').click(webThree);
            $('#webFour-button').click(webFour);
            $('#webFive-button').click(webFive);
        });
    };



    function webOne() {
        document.getElementById("websiteEmbed").src = one;
    }
    function webTwo() {
        document.getElementById("websiteEmbed").src = two;
    }
    function webThree() {
        document.getElementById("websiteEmbed").src = three;
    }
    function webFour() {
        document.getElementById("websiteEmbed").src = four;
    }
    function webFive() {
        document.getElementById("websiteEmbed").src = five;
    }
    function displaySelectedText() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    showNotification('Error:', result.error.message);
                }
            });
    }
    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
