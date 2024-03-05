def sharepoint_url_pruning(site_str):
    '''Isolates the site from the sharepoint url.'''
    
    iterator = iter(site_str.split('/'))
    res = []
    for item in iterator:
        res.append(item)
        if item == "sites" or item == "personal": 
            res.append(next(iterator, None))
            break
    
    return '/'.join(res)
           