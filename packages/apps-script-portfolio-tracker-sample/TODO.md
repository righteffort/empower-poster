* fill in Asset Classes automatically ... roughly this but need to set
  Name to Category when Name is missing (Cash) ... though we do that
  elsewhere anyay: 
```javascript
[...new Set(Object.values(classifications).flatMap(classifications =>
  classifications.map(c => c.classes.reverse())))]`
```
