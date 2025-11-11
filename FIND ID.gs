function myFunction() {
  var taskLists = Tasks.Tasklists.list().getItems();
  var res = taskLists.map(function(list) {
    return {
      taskListId: list.getId(),
      listName: list.getTitle(),
    };
  });
  Logger.log(res)
}