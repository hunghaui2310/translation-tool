from pptx import Presentation
import requests

if __name__ == '__main__':
    prs = Presentation('gg.pptx')


    def search_and_replace(input, output):
        """"search and replace text in PowerPoint while preserving formatting"""
        # Useful Links ;)
        # https://stackoverflow.com/questions/37924808/python-pptx-power-point-find-and-replace-text-ctrl-h
        # https://stackoverflow.com/questions/45247042/how-to-keep-original-text-formatting-of-text-with-python-powerpoint
        from pptx import Presentation
        prs = Presentation(input)
        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    text_frame = shape.text_frame
                    cur_text = text_frame.paragraphs[0].runs[0].text
                    new_text = call_api(cur_text)
                    print(new_text)
                    text_frame.paragraphs[0].runs[0].text = new_text
        prs.save(output)

    def call_api(text_translate):
        query = {'sl': 'auto', 'tl': 'vi', 'dt': 't', 'client': 'gtx', 'q': text_translate}
        response = requests.get("https://translate.googleapis.com/translate_a/single", query)
        return response.json()[0][0][0]


    search_and_replace('gg.pptx', 'ggOut.pptx')
