<?php
declare(strict_types=1);

namespace App\Shared\Support;

final class PhpView
{
    public function render(string $templatePath, array $data = []): void
    {
        if (!is_file($templatePath)) {
            throw new \RuntimeException('No se encontro la plantilla PHP: ' . $templatePath);
        }

        extract($data, EXTR_SKIP);
        require $templatePath;
    }
}
