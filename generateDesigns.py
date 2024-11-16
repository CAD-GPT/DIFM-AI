m = model.to(device)
context = torch.zeros((1, 1), dtype=torch.long, device=device)
shape1 = m.generate(context, max_new_tokens=2000)[0].tolist()
print(len(shape1))

print(decode(m.generate(context, max_new_tokens=2000)[0].tolist()))