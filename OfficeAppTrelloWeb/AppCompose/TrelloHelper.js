var ConnectToTrello = function () {
    console.log("Authenticating");
    Trello.authorize({
        type: 'popup',
        name: 'Outlook Trello Add-In',
        scope: { read: true, write: true, account: true },
        success: authenticationSuccess,
        error: authenticationError
    });
};

var GetTrelloLists = function () {
    console.log("Updating list of lists.");

    $('#trelloLists').empty();
    $('#trelloLists').append("<option value='false'>Select a List</option>");



    boardId = $('#trelloBoards').val();
    if (boardId == "false") {
        return;
    }



    Trello.get('/boards/' + boardId + '/lists',
        function (success) {

            for (var i = success.length - 1; i >= 0; i--) {
                list = success[i];
                var listOption = document.createElement("option");
                listOption.value = list.id;
                listOption.innerText = list.name;
                $('#trelloLists').append(listOption);
            };
        },
        function (error) {
            console.log(error);
            app.showNotification("Error getting lists", error);

        }
    );
};

var GetTrelloCards = function () {
    console.log("Updating list of cards.");
    $('#trelloCards').empty();

    listId = $('#trelloLists').val();

    if (listId == "false") {
        return;
    }


    Trello.get('/lists/' + listId + '/cards',
        function (success) {


            for (var i = success.length - 1; i >= 0; i--) {
                card = success[i];

                var cardHTML = "<div class=\"panel panel-default trellocard\" onclick=\"InsertCard('" + card.id + "')\"><div class=\"panel-body\">"+card.name+"</div></div>";
                $('#trelloCards').append(cardHTML);


            };
            $("#trelloCards").show();
        },
        function (error) {

            app.showNotification("Error getting cards", error);

        }
    );
};

var InsertCard = function (cardId) {
    Trello.get('/cards/' + cardId,
			function (card) {

			    Office.cast.item.toItemCompose(Office.context.mailbox.item).subject.setAsync("RE: " + card.name);
			    Office.context.mailbox.item.body.setSelectedDataAsync("<div class=\"border-left-width: 2px;border-left-color: #0067A3;border-left-style: solid;padding-left: 10px;\"><h2>" + card.name + "</h2><div>" + card.desc + "</div></div>", { coercionType: Office.CoercionType.Html });

			    app.showNotification(card.name, "Trello card added to message!");



			    //$('#cardName').text(success.name + "(" + cardId + ")");
			    //$('#cardDescription').text(success.desc);
			    //$('#cardLink').attr("href", success.shortUrl);
			},
			function (error) {
			    app.showNotification("Error getting card", error);

			}
		);
};

var authenticationSuccess = function () {
    $('#trelloConnect').hide();

    // Load Boards
    Trello.get('/member/me/boards',
		function (success) {
		    for (var i = success.length - 1; i >= 0; i--) {
		        board = success[i];
		        var boardOption = document.createElement("option");
		        boardOption.value = board.id;
		        boardOption.innerText = board.name;
		        $('#trelloBoards').append(boardOption);
		    };
		    $("#trelloInfoPanel").show();
		},
		function (error) {
		    console.log(error);
		    app.showNotification("Error loading boards", error);
		}
	);
}

var authenticationError = function (error) {
    console.log(error);
    app.showNotification("Authentication error", error);

}

