$(document).ready(function(){
    var selected = null;

    function removeHighlight()
    {
        $("#generatedName").removeClass("selected");
        $("#generatedTitle").removeClass("selected");
        $("#generatedPhone").removeClass("selected");

    }

    // Change Events

    $('.fontSize').on('change', function() {
        if(selected != null)
        {
            selected.css("font-size", this.value);
        }
    })

    $('.fontFamily').on('change', function() {
        if(selected != null)
        {
            selected.css("font-family", this.value);
        }
    })

    // Click Events

    $('.normal').click( function(){
        if(selected != null)
        {
            selected.css("font-weight", "normal");
        }
    });

    $('.bold').click( function(){
        if(selected != null)
        {
            selected.css("font-weight", "bold");
        }
    });

    $('.italic').click( function(){
        if(selected != null)
        {
            if(selected.css("font-style") == "italic")
                selected.css("font-style", "normal");
            else
                selected.css("font-style", "italic");
        }
    });

    $("#generatedName").click(function(){
        if($("#generatedName").hasClass("selected"))
        {
            removeHighlight();
            selected = null;
        }
        else
        {
            removeHighlight();
            $("#generatedName").addClass("selected");
            selected = $("#generatedName");
        }
    });

    $("#generatedTitle").click(function(){
        if($("#generatedTitle").hasClass("selected"))
        {
            removeHighlight();
            selected = null;
        }
        else
        {
            removeHighlight();
            $("#generatedTitle").addClass("selected");
            selected = $("#generatedTitle");
        }
    });

    $("#generatedPhone").click(function(){
        if($("#generatedPhone").hasClass("selected"))
        {
            removeHighlight();
            selected = null;
        }
        else
        {
            removeHighlight();
            $("#generatedPhone").addClass("selected");
            selected = $("#generatedPhone");
        }
    });

		
    $("#send").click(function(){
        $("#generatedName").text($("#name").val());
        $("#generatedTitle").text($("#title").val());
        var address = $("#address").val();
        $("#generatedAddress").text(address);
        $("#generatedAddress").attr("href", "http://maps.google.com/?q=" +
           encodeURI(address));
        var website = $("#website").val();
        $("#generatedWebsite").text(website);
        $("#generatedWebsite").attr("href", "http://" + website);
        $("#generatedPhone").text("Phone: " + $("#phone").val());


        var file = document.querySelector("input[type=file]").files[0];
        var img = document.querySelector("#generatedImg");
        var reader = new FileReader();

        reader.addEventListener("load", function(){
            img.src = reader.result;
        }, false);

        if(file)
            reader.readAsDataURL(file);
    });

});
