# coding: utf-8

import docx
from docx.oxml.shared import qn

from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl


def start_search(element, title):
    result = []

    if element.tag == qn('w:bookmarkStart'):
        if element.get(qn('w:name')) == title:
            result.append(element.get(qn('w:name')))
            result.append(element.get(qn('w:id')))

    return result


def end_search(element, bookmark):
    result = []

    if element.tag == qn('w:bookmarkEnd'):
        if element.get(qn('w:id')) == bookmark[1]:
            result.append('Is found')

    if element.tag == qn('w:bookmarkStart'):
        result.append('w:bookmarkStart')

    return result


def delete_bookmark(file_name,
                    bookmark_name):
    doc = docx.Document(file_name)
    body = doc.part.element.body

    run_end_search = False
    run_start_search = True
    search_bookmark = []

    for any_element in body:
        if isinstance(any_element, CT_P) or isinstance(any_element, CT_Tbl):

            for elem in any_element:
                if run_start_search:
                    search_bookmark = start_search(elem, bookmark_name)

                    if search_bookmark:
                        any_element.remove(elem)
                        run_end_search = True
                        run_start_search = False
                        continue

                if run_end_search:
                    search_end = end_search(elem, search_bookmark)

                    if 'w:bookmarkStart' in search_end:
                        continue
                    else:
                        any_element.remove(elem)

                    if search_end:
                        doc.save('testDelete.docx')
                        return

        child = any_element.getchildren()
        if not child:
            body.remove(any_element)


def find_begin_bookmark(any_element,
                  bookmark_name):
    search_bookmark = []

    for elem in any_element:
        search_bookmark = start_search(elem, bookmark_name)

        if not search_bookmark:
            any_element.remove(elem)
        else:
            break

    return search_bookmark


def find_end_bookmark(any_element,
                      search_bookmrk):
    is_found = False

    for elem in any_element:
        book_end = end_search(elem, search_bookmrk)

        if 'Is found' in book_end:
            is_found = True
            continue

        if is_found:
            any_element.remove(elem)

    return is_found


def copy_bookmark(file_name,
                  bookmark_name):
    in_doc = docx.Document(file_name)
    in_body = in_doc.part.element.body

    run_end_search = False
    run_begin_search = True
    run_remove_other = False
    search_bookmrk = []

    for any_element in in_body:
        if isinstance(any_element, CT_P) or isinstance(any_element, CT_Tbl):
            if run_begin_search:
                search_bookmrk = find_begin_bookmark(any_element, bookmark_name)

                if not search_bookmrk:
                    in_body.remove(any_element)
                else:
                    run_begin_search = False
                    run_end_search = True
                    continue

            if run_end_search:
                is_found_end = find_end_bookmark(any_element, search_bookmrk)

                if not is_found_end:
                    continue
                else:
                    run_end_search = False
                    run_remove_other = True
                    continue

            if run_remove_other:
                in_body.remove(any_element)

    if not run_begin_search:
        in_doc.save('testcopy.docx')


if __name__ == '__main__':
    delete_bookmark('20171214-IC-template.docx', 'ActivateRN')
    copy_bookmark('20171214-IC-template.docx', 'test')