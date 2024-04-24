//Реализовать консольное приложение со следующей логикой: запуск трех(3) потоков, выполняющихся асинхронно или параллельно. 
//Каждый поток реализует доступ к шареному ресурсу, а именно инстансу Microsoft Word Application. 
//При этом каждый поток выполняет свой делегат над шаренным ресурсом, описывающий что нужно делать каждому потоку в отдельности. 
//При этом, хоть потоки и запускаются асинхронно параллельно, но очередь их доступа к шаренному ресурсу (по факту к началу выполнения связанного с ним делегата) 
//должна описываться и настраиваться в координаторе: в каждый конкретный момент времени только один поток будет осуществлять работу с шаренным объектом. 
//К примеру: делегат первого потока, выполняющийся первым, создает пустой документ. Последний, третий, поток выполняет делегат сохранения документа в файл. 
//Второй по очередности поток может наполнить документ, а может спать некоторое время. 
//Резюмируя: запустить три параллельно запущенных потока, работающих формально в последовательной синхронной модели над шаренным ресурсом по очередности, заданной в координаторе. 
//При отсутствии Windows или нежелании устанавливать Word шаренный ресурс может быть заменен на LibreOffice и его SDK, 
//либо на прокси объект, имеющий схожий интерфейс с реальным объектом.

using Microsoft.Office.Interop.Word;

var documentHandler = new DocumentHandler();
var actions = new List<Action> { documentHandler.CreateDocument, documentHandler.AddContent, documentHandler.SaveDocument };
var threadCoordinator = new ThreadCoordinator(actions, new List<int> { 0, 1, 2 });
threadCoordinator.StartThreads();

//тест с выводом на консоль

//var actions = new List<Action> { Print1, Print2, Print3 };
//var threadCoordinator = new ThreadCoordinator(actions, new List<int> { 2, 0, 1 });
//threadCoordinator.StartThreads();

//void Print1()
//{
//    Thread.Sleep(2000);
//    Console.WriteLine(1);
//}

//void Print2()
//{
//    Thread.Sleep(2000);
//    Console.WriteLine(2);
//}

//void Print3()
//{
//    Thread.Sleep(2000);
//    Console.WriteLine(3);
//}

public class ThreadCoordinator
{
    private List<Action> _actions;
    private List<int> _queue;
    private int _currentQueueIndex;
    
    public ThreadCoordinator(List<Action> actions, List<int> queue)
    {
        _actions = actions;
        _queue = queue;
        _currentQueueIndex = 0;
    }

    public void StartThreads()
    {
        for (var i = 0; i < _actions.Count; i++) 
        {
            var thread = new Thread(ExecuteAction);
            thread.Start(i);
        }
    }

    private void ExecuteAction(object i)
    {
        var index = int.Parse(i.ToString());

        while (index != _queue[_currentQueueIndex])
            Thread.Sleep(500);

        _actions[index]();

        _currentQueueIndex++;
    }
}

public class DocumentHandler
{
    private Document _document;
    public void CreateDocument() => _document = new Application().Documents.Add();

    public void AddContent() => _document.Content.Text += "Some text";

    public void SaveDocument() => _document.SaveAs("Test.doc");
}
