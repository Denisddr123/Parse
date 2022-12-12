import org.apache.hc.client5.http.classic.methods.HttpGet;
import org.apache.hc.client5.http.impl.classic.CloseableHttpClient;
import org.apache.hc.client5.http.impl.classic.CloseableHttpResponse;
import org.apache.hc.client5.http.impl.classic.HttpClients;
import org.apache.hc.core5.http.HttpEntity;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.xmlbeans.XmlException;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.jsoup.select.Evaluator;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

public class Parse {
    static Elements thisElements;
    static NewDocument newDocument;
    static Document document;

    public static void main(String[] args) throws IOException, InterruptedException, InvalidFormatException {
        try (CloseableHttpClient httpClient = HttpClients.createDefault()) {

            HttpGet httpGet = new HttpGet("https://rsport.ria.ru/football/");
            try (CloseableHttpResponse response = httpClient.execute(httpGet)) {
                System.out.println(response.getCode() + " " + response.getReasonPhrase());
                HttpEntity entity = response.getEntity();

                document = Jsoup.parse(entity.getContent(), null, "https://rsport.ria.ru/football/");
                entity.close();
            }
            Pattern pattern = Pattern.compile("[\\w\\W]+");
            Elements elements = document.getElementsByAttributeValueMatching("data-title", pattern);
            System.out.println("Вывод всех ссылок на статьи.");
            for (Element element: elements) {
                System.out.println("Ссылка "+element.attr("data-url")+", тема:"+element.attr("data-title"));
            }

            Pattern finalPattern = Pattern.compile("Спартак|Краснодар|Зенит|ЦСКА|Динамо|Локомотив", Pattern.CASE_INSENSITIVE);
            List<Element> list = elements.stream().map(element0 -> element0.getElementsByAttributeValueMatching("data-title", finalPattern).first()).filter(Objects::nonNull).collect(Collectors.toList());
            Elements elements2 = new Elements(list);
            thisElements = elements2;
            System.out.println("Вывод ссылок на статьи включающие слова Спартак|Краснодар|Зенит|ЦСКА|Динамо|Локомотив");
            for (Element element2: elements2) {
                System.out.println("Ссылка "+element2.attr("data-url")+", тема:"+element2.attr("data-title"));
            }


            if (thisElements != null && thisElements.size()!= 0) {
                httpGet = new HttpGet(Objects.requireNonNull(thisElements.first()).attr("data-url"));
                newDocument = new NewDocument();
                try (CloseableHttpResponse response = httpClient.execute(httpGet)) {
                    HttpEntity entity = response.getEntity();

                    document = Jsoup.parse(entity.getContent(), null, "https://rsport.ria.ru/");
                    entity.close();
                }

                System.out.println("Создание файла create.docx с текстом из статьи");
                elements = document.getElementsByTag("h1");
                for (Element el:
                        elements) {
                    h1Element(el);
                }
                Element elements1 = document.select("div.article__announce").first();
                if (elements1 != null) {
                    boolean boolTag = new Evaluator.Tag("img").matches(elements1, new Element("img"));
                    if (boolTag) {
                        elements1 = elements1.getElementsByTag("img").first();
                        mediaElement(elements1);
                    }
                }

                Element element2 = document.select("div.article__body").first();

                elements2 = Objects.requireNonNull(element2).select("div.article__block");
                for (Element el:
                        elements2) {
                    switchElement(el);
                }
            }
            newDocument.writeDocument();
        } catch (XmlException e) {
            throw new RuntimeException(e);
        }
    }

    static void switchElement (Element element) throws IOException {
        String str = element.attr("data-type");
        int numId = 4;
        switch (str) {
            case "h1": h1Element(element);
            break;
            case "h2": h2Element(element);
            break;
            case "list": listElement(element, numId);
            break;
            case "media": mediaElement(element);
            break;
            case "text": textElement(element);
            break;
            case "quote": quoteElement(element);
            break;
            case "h3": h3Element(element);
            break;
            case "article": articleElement(element);
            break;
            case "table": tableElement(element);
            break;
            default: break;
        }
    }

    static void h1Element (Element element) {
        newDocument.addH1(element.text());
    }
    static void h2Element (Element element) {
        newDocument.addH2(element.text());
    }
    static void h3Element (Element element) {
        newDocument.addH3(element.text());
    }

    static void listElement (Element element, int numId) {
        ArrayList<String> list = new ArrayList<>();
        Elements elements = element.getElementsByTag("li");
        for (Element el:
             elements) {
            list.add(el.text());
        }
        if (elements.size() == 1) {
            newDocument.addEnumeration(list, numId);
        } else {
            newDocument.addEnumeration(list, numId, 0, true);
        }
    }

    static void mediaElement (Element element) throws IOException {
        element = element.getElementsByTag("img").first();
        newDocument.addImage(Objects.requireNonNull(element).attr("src"));
    }
    static void textElement (Element element) {
        newDocument.addText(element.text());
    }
    static void quoteElement (Element element) {
        newDocument.addQuote(element.text());
    }
    static void articleElement (Element element) throws IOException {
        Element element1 = element.getElementsByTag("img").first();
        String str = Objects.requireNonNull(element1).attr("src");
        if (!str.startsWith("http")) {
            str = Objects.requireNonNull(element1).attr("data-src");
        }
        newDocument.addArticle(str, element.text());
    }
    static void tableElement (Element element) {
        Element element1 = element.getElementsByTag("thead").first();
        Elements elements = Objects.requireNonNull(element1).getElementsByTag("td");
        Element element2 = element.getElementsByTag("tbody").first();
        Elements elements1 = Objects.requireNonNull(element2).getElementsByTag("tr");
        String[][] arr = new String[elements1.size()+1][elements.size()];
        int i=0, j=1;

        for (Element el:
             elements) {
            arr[0][i++] = el.text();
        }

        for (Element el:
             elements1) {
            Elements elements2 = el.getElementsByTag("td");
            i=0;
            for (Element el2:
                 elements2) {
                arr[j][i++] = el2.text();
            }
            j++;
        }
        newDocument.addTable(arr);
    }
}
