<?php
declare(strict_types=1);

namespace App\Cxp\ServiciosMarcas\Domain;

final class HistoryLabelResolver
{
    /**
     * @var array<string, array{label:string, prefix:string}>
     */
    private array $outputConfig;

    /**
     * @param array<string, array{label:string, prefix:string}> $outputConfig
     */
    public function __construct(array $outputConfig)
    {
        $this->outputConfig = $outputConfig;
    }

    public function resolveLabel(string $fileName): string
    {
        $brandKey = $this->detectBrandKey($fileName);
        if ($brandKey === null) {
            return 'SERVICIOS';
        }

        return (string)($this->outputConfig[$brandKey]['label'] ?? strtoupper($brandKey));
    }

    public function detectBrandKey(string $fileName): ?string
    {
        $name = strtolower($fileName);
        foreach ($this->outputConfig as $key => $config) {
            $prefix = strtolower((string)($config['prefix'] ?? ''));
            if ($prefix !== '' && str_starts_with($name, $prefix)) {
                return (string)$key;
            }
        }

        return null;
    }

    /**
     * @return string[]
     */
    public function brandOrder(): array
    {
        return array_keys($this->outputConfig);
    }
}
